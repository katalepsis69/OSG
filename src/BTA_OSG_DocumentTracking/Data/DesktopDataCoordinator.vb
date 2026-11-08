Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Data

Namespace BTA_OSG
    ''' <summary>
    ''' Coordinates data dispatch between enterprise SQL Server backend services
    ''' and the local embedded database cache, ensuring complete offline capability
    ''' while syncing transparently to SQL Server whenever connected.
    ''' </summary>
    Partial Public Class DesktopDataCoordinator
        ' Live reachability, not the startup probe. The poller refreshes it on every pull and
        ' a failed write drops it immediately, so an outage costs one connect timeout instead
        ' of one per action for the rest of the session.
        Private _sqlAvailable As Boolean

        Public ReadOnly Property IsDatabaseConnected As Boolean
            Get
                Return _sqlAvailable
            End Get
        End Property

        Public Sub New(isDatabaseConnected As Boolean)
            _sqlAvailable = isDatabaseConnected
        End Sub

        ''' <summary>
        ''' Registers a document. Connected: SQL Server owns the code, the id, and the
        ''' workflow rows, and the cache is refreshed from SQL. Offline, or after a failed SQL
        ''' write: the cache owns the row, exactly as before.
        ''' Returns the authoritative DocCode so the operator sees the code that was filed.
        ''' </summary>
        Public Function RegisterDocument(code As String, docType As String, title As String, origin As String, dest As String, cab As String, shelf As String, box As String, gDriveUrl As String, initialStatus As String, assignedStaff As String, flowDir As String, assignedSec As String, deadline As String, punchlist As String, lastAction As String, registeredByName As String, registeredByUserId As Integer, Optional externalControlNumber As String = "") As String
            If IsDatabaseConnected AndAlso AppStartup.DocService IsNot Nothing Then
                Dim fields As New RegistrationInput With {
                    .DocType = docType,
                    .Title = title,
                    .Origin = origin,
                    .Destination = dest,
                    .Cabinet = cab,
                    .Shelf = shelf,
                    .Box = box,
                    .GDriveUrl = gDriveUrl,
                    .InitialStatus = initialStatus,
                    .AssignedStaff = assignedStaff,
                    .FlowDirection = flowDir,
                    .AssignedSection = assignedSec,
                    .Deadline = deadline,
                    .LastAction = lastAction,
                    .RegisteredByUserId = registeredByUserId,
                    .ExternalControlNumber = externalControlNumber
                }
                Dim doc = RegisterDocumentInSqlInternal(fields)
                If doc IsNot Nothing Then Return doc.DocCode
            End If

            ' Offline, or the SQL write failed: the cache owns the row, as before.
            EmbeddedDB.AddDocument(code, docType, title, origin, dest, cab, shelf, box, gDriveUrl, initialStatus, assignedStaff, flowDir, assignedSec, deadline, punchlist, lastAction, externalControlNumber, isOffline:=True, createdByUserId:=registeredByUserId)
            EmbeddedDB.LogAudit(registeredByName, String.Format("Registered New OSG Document [{0}] : {1} (Auto-routed to {2})", code, title, assignedSec), isOffline:=True, userId:=registeredByUserId)
            Return code
        End Function

        ''' <summary>
        ''' The connected registration. The code comes from tbl_DocumentSequences under UPDLOCK
        ''' (Services/DocumentService.vb:45), never from the local row counter, so two
        ''' workstations cannot mint the same DocCode. The document id is SQL's IDENTITY and
        ''' the cache row is a mirror of it, which is what stops one document from existing
        ''' under two keys. Returns Nothing on failure so the caller falls back to the cache.
        ''' </summary>
        Private Function RegisterDocumentInSqlInternal(fields As RegistrationInput) As Document
            Try
                Dim targetDt As DateTime? = Nothing
                Dim parsedDt As DateTime
                If DateTime.TryParse(fields.Deadline, parsedDt) Then targetDt = parsedDt

                ' The cache row's DateReceived is when the document actually arrived; without
                ' carrying it, replay stamps the sync moment and offline registrations are
                ' measured from the wrong day in analytics and on the printed slip.
                Dim receivedDt As DateTime? = Nothing
                If Not String.IsNullOrWhiteSpace(fields.DateReceived) AndAlso DateTime.TryParse(fields.DateReceived, parsedDt) Then
                    receivedDt = parsedDt
                End If

                Dim doc As Document = Nothing
                ' One connection, one transaction: the sequence reservation, the document row,
                ' the workflow update, the assignment, the routing log, and the storage
                ' movement either land whole or not at all, so a mid-sequence failure cannot
                ' strand an orphan SQL document for the outbox to re-register under a fresh code.
                Using conn = AppStartup.ConnectionFactory.CreateConnection()
                    doc = WriteRegistrationWorkflow(fields, targetDt, receivedDt, conn)
                End Using
                If doc Is Nothing Then Return Nothing

                ' Refresh the operator's grid in the same click via quick LAN read. This runs
                ' after the commit: the SQL write is already durable, so a refresh failure
                ' must not turn a registered document into a local re-registration.
                Try
                    EmbeddedDB.ApplySnapshots(BuildSqlSnapshot("Documents", "RoutingLogs", "Movements", "AuditTrail"))
                Catch refreshEx As Exception
                    System.Diagnostics.Trace.WriteLine("Post-registration cache refresh failed: " & refreshEx.Message)
                End Try
                Return doc
            Catch ex As Exception
                ' The fallback itself must reach the operator: a silent drop here looks like a
                ' successful registration while the row only ever lived in the local cache.
                _sqlAvailable = False
                System.Diagnostics.Trace.WriteLine("SQL Server Document Register error (falling back to embedded cache): " & ex.Message)
                Program.ReportOperatorWarning("Registration was saved to this workstation's offline store because the office server refused the write (" & ex.Message & "). It retries automatically; if this repeats, run Station Setup.")
                Return Nothing
            End Try
        End Function

        ''' <summary>The registration transaction body: the six writes (sequence-backed
        ''' document row, workflow state, assignment, routing log, storage movement,
        ''' storage location pointer) land whole or not at all. Returns Nothing when the
        ''' service refused to mint the document; the caller must not fall back to the
        ''' cache in that case either (Nothing means refused, not failed).</summary>
        Private Function WriteRegistrationWorkflow(fields As RegistrationInput, targetDt As DateTime?, receivedDt As DateTime?, conn As Microsoft.Data.SqlClient.SqlConnection) As Document
            Using tx = conn.BeginTransaction()
                Try
                    Dim doc = AppStartup.DocService.RegisterDocumentWithWorkflow(fields.Title, TypeCodeFor(fields.DocType), fields.FlowDirection, fields.Origin, fields.Destination, targetDt, fields.GDriveUrl, fields.LastAction, fields.RegisteredByUserId, transaction:=tx, receivedDate:=receivedDt, externalControlNumber:=fields.ExternalControlNumber, preferredDocCode:=fields.PreferredDocCode, requesterGender:=fields.RequesterGender)
                    If doc Is Nothing Then Return Nothing

                    ' The desktop registers INTO a workflow: status, desk, and owner are part of the
                    ' record, not defaults to be discovered later.
                    Dim status = AppStartup.ReferenceDataRepo.GetStatusByCode(If(String.IsNullOrWhiteSpace(fields.InitialStatus), "RECEIVED", fields.InitialStatus))
                    If status Is Nothing Then status = AppStartup.ReferenceDataRepo.GetStatusByCode("RECEIVED")
                    Dim statusId As Integer = If(status IsNot Nothing, status.StatusID, 1)
                    Dim effectiveSection = If(String.IsNullOrWhiteSpace(fields.AssignedSection), DocumentService.GetDefaultSectionForCategory(fields.DocType), fields.AssignedSection)
                    AppStartup.DocumentRepo.UpdateWorkflowState(doc.DocumentID, statusId, effectiveSection, Nothing, fields.LastAction, fields.RegisteredByUserId, transaction:=tx)

                    Dim ownerId = ResolveUserId(fields.AssignedStaff)
                    If ownerId.HasValue Then
                        AppStartup.DocumentRepo.AddAssignment(New DocumentAssignment With {
                            .DocumentID = doc.DocumentID,
                            .AssignedUserID = ownerId,
                            .AssignedByUserID = fields.RegisteredByUserId,
                            .AssignedAtUTC = DateTime.UtcNow,
                            .Remarks = "Registered and routed to " & effectiveSection
                        }, tx)
                    End If

                    AppStartup.RoutingRepo.Insert(New RoutingLog With {
                        .DocumentID = doc.DocumentID,
                        .FromStatusID = Nothing,
                        .ToStatusID = statusId,
                        .FromOffice = fields.Origin,
                        .ToOffice = fields.Destination,
                        .RoutedByUserID = fields.RegisteredByUserId,
                        .RoutedAtUTC = DateTime.UtcNow,
                        .RoutingRemarks = "Document registered in OSG System."
                    }, tx)

                    Dim locationId = AppStartup.StorageRepo.GetOrCreateByKey(fields.Cabinet, fields.Shelf, fields.Box, tx)
                    AppStartup.StorageRepo.InsertMovement(New DocumentMovement With {
                        .DocumentID = doc.DocumentID,
                        .StorageLocationID = locationId,
                        .MovedByUserID = fields.RegisteredByUserId,
                        .MovedAtUTC = DateTime.UtcNow,
                        .MovementReason = "Initial storage assignment."
                    }, tx)
                    AppStartup.DocumentRepo.UpdateStorageLocation(doc.DocumentID, locationId, tx)

                    tx.Commit()
                    Return doc
                Catch
                    tx.Rollback()
                    Throw
                End Try
            End Using
        End Function

        ''' <summary>
        ''' The connected half of user enrolment: upsert into tbl_Users, issue the badge,
        ''' and assign the role, all in one transaction so a mid-way failure cannot leave
        ''' a user with no card or no role. Shared by RegisterUser and ReplayPendingUsers.
        ''' claimGuard (first-administrator claim only) refuses a second SYSADMIN inside the
        ''' transaction, because the claim flow's UI-level check cannot stop two seats that
        ''' claim within the same window or replay the same claim from two caches.
        ''' Returns False so the caller keeps the row pending instead of losing it.
        ''' </summary>
        Private Function EnsureUserInSql(fullName As String, role As String, office As String, cardUid As String, canRoute As Boolean, canMove As Boolean, canSoftCopy As Boolean, enrolledByUserId As Integer, Optional claimGuard As Boolean = False) As Boolean
            Dim cleanUid = If(cardUid, "").Trim().ToUpperInvariant()
            Try
                Using conn = AppStartup.ConnectionFactory.CreateConnection()
                    Using tx = conn.BeginTransaction()
                        Try
                            Dim existing = AppStartup.UserRepo.GetByCardPublicID(cleanUid)
                            If existing IsNot Nothing Then
                                existing.FullName = fullName
                                existing.Office = office
                                existing.IsActive = True
                                existing.CanRoute = canRoute
                                existing.CanMove = canMove
                                existing.CanSoftCopy = canSoftCopy
                                AppStartup.UserRepo.Update(existing, enrolledByUserId, tx)
                                AppStartup.UserRepo.AssignRole(existing.UserID, RoleCodeFor(role), enrolledByUserId, tx)
                            Else
                                InsertUserWithCard(fullName, role, office, cleanUid, canRoute, canMove, canSoftCopy, enrolledByUserId, tx, claimGuard)
                            End If
                            tx.Commit()
                            Return True
                        Catch
                            tx.Rollback()
                            Throw
                        End Try
                    End Using
                End Using
            Catch refusal As GuardRefusedException
                ' A refused claim is a decision, not a dead host: the caller (the claim
                ' dialog) must show it, and the row must not be cached for a replay that
                ' would be refused by the same guard anyway.
                Throw
            Catch ex As Exception
                System.Diagnostics.Trace.WriteLine("SQL Server user enrolment error: " & ex.Message)
                Return False
            End Try
        End Function

        ''' <summary>The new-user enrolment contract, inside the caller's transaction: the
        ''' claim guard (only one SYSADMIN may ever be claimed), the insert with the
        ''' UNIQUE-username badge-suffix retry, the badge issue, and the role assign.
        ''' Any guard refusal or SQL failure rethrows so the caller rolls back whole.</summary>
        Private Sub InsertUserWithCard(fullName As String, role As String, office As String, cleanUid As String, canRoute As Boolean, canMove As Boolean, canSoftCopy As Boolean, enrolledByUserId As Integer, tx As Microsoft.Data.SqlClient.SqlTransaction, claimGuard As Boolean)
            If claimGuard AndAlso RoleCodeFor(role) = RbacPolicy.ROLE_SYSADMIN AndAlso
               AppStartup.UserRepo.HasAnyUserWithRole(RbacPolicy.ROLE_SYSADMIN) Then
                Throw New GuardRefusedException("An administrator has already been claimed on this server.")
            End If
            Dim newUser As New User With {
                .Username = SanitizeUsername(fullName),
                .FullName = fullName,
                .Office = office,
                .IsActive = True,
                .CanRoute = canRoute,
                .CanMove = canMove,
                .CanSoftCopy = canSoftCopy
            }
            Dim userId As Integer
            Try
                userId = AppStartup.UserRepo.Insert(newUser, enrolledByUserId, tx)
            Catch dup As Microsoft.Data.SqlClient.SqlException
                ' Username is UNIQUE too; two staff with the same name get the badge suffix.
                newUser.Username = SanitizeUsername(fullName) & "_" & Right(cleanUid, 4).ToLowerInvariant()
                userId = AppStartup.UserRepo.Insert(newUser, enrolledByUserId, tx)
            End Try

            AppStartup.CardService.IssueCard(userId, cleanUid, "Enrolled from desktop Admin", enrolledByUserId, tx)
            AppStartup.UserRepo.AssignRole(userId, RoleCodeFor(role), enrolledByUserId, tx)
        End Sub

        ''' <summary>
        ''' One document registration handed to the SQL path, either from the registry form
        ''' or replayed from an offline cache row. Plain field bag: both callers read from
        ''' positional sources, so a constructor would just move the same arity around.
        ''' </summary>
        Private Class RegistrationInput
            Public Property DocType As String
            Public Property Title As String
            Public Property Origin As String
            Public Property Destination As String
            Public Property Cabinet As String
            Public Property Shelf As String
            Public Property Box As String
            Public Property GDriveUrl As String
            Public Property InitialStatus As String
            Public Property AssignedStaff As String
            Public Property FlowDirection As String
            Public Property AssignedSection As String
            Public Property Deadline As String
            ' When the document actually arrived, carried from the cache row so a replayed
            ' offline registration keeps its original received time instead of the sync time.
            Public Property DateReceived As String
            Public Property LastAction As String
            Public Property RegisteredByUserId As Integer
            Public Property ExternalControlNumber As String
            ' Portal-captured requester gender (db/scripts 016). Without carrying it, a
            ' replayed offline import silently drops the column and gender analytics
            ' undercount on seats that imported while SQL was down.
            Public Property RequesterGender As String
            ' The code minted offline. SQL reuses it when free, falls back to the shared
            ' sequence on a collision, so the operator's offline paperwork stays valid.
            Public Property PreferredDocCode As String
        End Class

        ''' <summary>
        ''' Display name to document type code, matching the codes seeded by
        ''' db/scripts/008_osg_target_migration.sql:5-12. The SQL service resolves the prefix
        ''' from the code, so this map is the whole contract between the two vocabularies.
        ''' </summary>
        Private Shared Function TypeCodeFor(docType As String) As String
            Select Case If(docType, "").Trim().ToUpperInvariant()
                Case "LEGISLATIVE", "LEG" : Return "LEG"
                Case "FINANCE", "FIN" : Return "FIN"
                Case "TRAVEL", "TRAVEL ORDER", "TO" : Return "TRAVEL"
                Case Else : Return "REG_COMM"
            End Select
        End Function

        ''' <summary>
        ''' Staff full name to SQL user id, for the assignment row. The staff combo carries
        ''' names, and tbl_Users is small, so one read beats a new indexed lookup.
        ''' </summary>
        Friend Shared Function ResolveUserId(fullName As String) As Integer?
            If String.IsNullOrWhiteSpace(fullName) OrElse AppStartup.UserRepo Is Nothing Then Return Nothing
            For Each u As User In AppStartup.UserRepo.GetAll()
                If String.Equals(u.FullName, fullName.Trim(), StringComparison.OrdinalIgnoreCase) Then Return u.UserID
            Next
            Return Nothing
        End Function

        ''' <summary>
        ''' Status id to seeded code, for the terminal-state guard. Unknown ids return ""
        ''' and the guard then stays open, exactly like the mirror's ISNULL fallback.
        ''' </summary>
        Friend Shared Function StatusCodeFor(statusId As Integer) As String
            If AppStartup.ReferenceDataRepo Is Nothing Then Return ""
            For Each s As DocumentStatus In AppStartup.ReferenceDataRepo.GetDocumentStatuses()
                If s.StatusID = statusId Then Return s.StatusCode
            Next
            Return ""
        End Function

        Public Sub RouteDocument(docId As Integer, fromOffice As String, toOffice As String, action As String, remarks As String, routedByName As String, routedByUserId As Integer)
            If IsDatabaseConnected AndAlso AppStartup.RoutingService IsNot Nothing Then
                ' Status update and custody log are one fact: a crash between them must not
                ' leave a moved document with no routing entry, so both join one transaction.
                If TryRunSql("RouteDocument", Sub() WriteRouteTransaction(docId, fromOffice, toOffice, action, remarks, routedByUserId)) Then
                    PullAfterWrite("Documents", "RoutingLogs", "AuditTrail")
                    Return
                End If
            End If
            EmbeddedDB.AddRoutingLog(docId, fromOffice, toOffice, routedByName, action, remarks)
            EmbeddedDB.LogAudit(routedByName, String.Format("Routed Doc #{0} from {1} to {2}: {3}", docId, fromOffice, toOffice, action))
        End Sub

        Public Sub ApplyDirective(docId As Integer, directiveText As String, assignedTo As String, notes As String, staffName As String, staffUserId As Integer)
            If IsDatabaseConnected AndAlso AppStartup.DirectiveService IsNot Nothing Then
                ' The UI combo carries the seeded directive names; resolve the type from the
                ' matching DirectiveCode so the status side effect (Approved & Archived) and
                ' the stored type match what the SQL schema defines.
                Dim directiveTypeId As Integer = 1
                For Each dt As DirectiveType In AppStartup.ReferenceDataRepo.GetDirectiveTypes()
                    If String.Equals(dt.DirectiveCode, DirectiveCodeFor(directiveText), StringComparison.OrdinalIgnoreCase) Then
                        directiveTypeId = dt.DirectiveTypeID
                        Exit For
                    End If
                Next
                If TryRunSql("ApplyDirective", Sub() AppStartup.DirectiveService.IssueDirective(docId, directiveTypeId, directiveText, notes, staffUserId, assignedTo)) Then
                    PullAfterWrite("Directives", "Documents", "AuditTrail")
                    Return
                End If
            End If
            EmbeddedDB.AddDirective(docId, directiveText, assignedTo, notes, staffName, directiveCode:=DirectiveCodeFor(directiveText))
            EmbeddedDB.LogAudit(staffName, String.Format("Applied SG Directive [{0}] to Doc ID #{1}", directiveText, docId))
        End Sub

        ''' <summary>
        ''' UI directive name to seeded tbl_DirectiveTypes.DirectiveCode (db/scripts/007).
        ''' Unmapped names land on IMMEDIATE_ACTION, the type with no status side effect.
        ''' </summary>
        Public Shared Function DirectiveCodeFor(directiveName As String) As String
            Select Case If(directiveName, "").Trim()
                Case "Referred to Committee on Rules", "Referred to Committee" : Return "REFER_COMMITTEE"
                Case "Forwarded for Speaker Signature" : Return "FORWARD_SPEAKER"
                Case "Under OSG Administrative Review" : Return "ADMIN_REVIEW"
                Case "Approved & Archived" : Return "APPROVE_ARCHIVE"
                Case Else : Return "IMMEDIATE_ACTION"
            End Select
        End Function

        Public Sub MoveStorage(docId As Integer, fromLoc As String, toLoc As String, reason As String, movedByName As String, movedByUserId As Integer)
            If IsDatabaseConnected AndAlso AppStartup.StorageService IsNot Nothing Then
                ' The whole point of a move is which box the folder went to, so resolve the real
                ' landmark first; MoveDocument used to be handed a hardcoded location id. The
                ' landmark creation and the movement join one transaction: a created-but-unmoved
                ' landmark is harmless, but a movement pointing at a landmark that failed to
                ' commit is not, so both commit together.
                Dim parts = If(toLoc, "").Split("|"c)
                Dim cab = If(parts.Length > 0, parts(0), "")
                Dim shelf = If(parts.Length > 1, parts(1), "")
                Dim box = If(parts.Length > 2, parts(2), "")
                If TryRunSql("MoveStorage", Sub() WriteMoveTransaction(docId, cab, shelf, box, reason, movedByUserId)) Then
                    PullAfterWrite("Documents", "Movements", "AuditTrail")
                    Return
                End If
            End If
            EmbeddedDB.AddMovementLog(docId, fromLoc, toLoc, movedByName, reason)
            EmbeddedDB.LogAudit(movedByName, String.Format("Transferred Doc #{0} storage location from {1} to {2}", docId, fromLoc, toLoc))
        End Sub

        ''' <summary>The connected routing transaction body: status update, custody log,
        ''' and audit row are one fact (a crash between them must not leave a moved
        ''' document with no routing entry). The terminal-state guard and the
        ''' non-regression rule throw here so TryRunSql sees the refusal exactly as
        ''' before the body was named.</summary>
        Private Sub WriteRouteTransaction(docId As Integer, fromOffice As String, toOffice As String, action As String, remarks As String, routedByUserId As Integer)
            Using conn = AppStartup.ConnectionFactory.CreateConnection()
                Using tx = conn.BeginTransaction()
                    Try
                        Dim currentDoc = AppStartup.DocumentRepo.GetById(docId)
                        If currentDoc Is Nothing Then Throw New InvalidOperationException("Document " & docId.ToString() & " no longer exists.")
                        Dim status = AppStartup.ReferenceDataRepo.GetStatusByCode(If(String.IsNullOrWhiteSpace(action), "ROUTED", action))

                        ' A filed, released, archived, or completed document has left the
                        ' active workflow: refuse the re-route instead of silently
                        ' reopening it with nothing but a routine custody log.
                        Dim currentCode = StatusCodeFor(currentDoc.StatusID)
                        If DocumentStatus.IsTerminalStatus(currentCode) AndAlso
                           (status Is Nothing OrElse Not String.Equals(status.StatusCode, currentCode, StringComparison.OrdinalIgnoreCase)) Then
                            Throw New GuardRefusedException("This document is already " & DocumentStatus.DisplayName(currentCode) & " and can no longer be routed.")
                        End If

                        ' An action that is not a status code must not regress the
                        ' document: keep its current status, exactly as the offline
                        ' and replay paths do. Only the destination and custody log
                        ' record the transmittal.
                        Dim statusId As Integer = If(status IsNot Nothing, status.StatusID, currentDoc.StatusID)

                        ' A destination office is not a desk: AssignedSection stays
                        ' (COALESCE via Nothing), matching the offline path, so a
                        ' document routed connected and one routed offline end up
                        ' visible at the same desk after sync.
                        AppStartup.DocumentRepo.UpdateWorkflowState(docId, statusId, Nothing, Nothing, "Routed to " & toOffice & ": " & action, routedByUserId, destinationOffice:=toOffice, originOffice:=fromOffice, transaction:=tx)

                        Dim log As New RoutingLog With {
                            .DocumentID = docId,
                            .FromStatusID = Nothing,
                            .ToStatusID = statusId,
                            .FromOffice = fromOffice,
                            .ToOffice = toOffice,
                            .RoutingRemarks = remarks,
                            .RoutedByUserID = routedByUserId,
                            .RoutedAtUTC = DateTime.UtcNow
                        }
                        AppStartup.RoutingRepo.Insert(log, tx)
                        If AppStartup.AuditService IsNot Nothing Then
                            AppStartup.AuditService.LogEvent("ROUTING_LOG_ADDED", "RoutingLog", docId.ToString(), Nothing, Nothing, Nothing, True, Nothing, tx)
                        End If
                        tx.Commit()
                    Catch
                        tx.Rollback()
                        Throw
                    End Try
                End Using
            End Using
        End Sub

        ''' <summary>The connected move transaction body: landmark creation and the
        ''' movement commit together, because a movement pointing at a landmark that
        ''' failed to commit is the one outcome the move must never produce.</summary>
        Private Sub WriteMoveTransaction(docId As Integer, cab As String, shelf As String, box As String, reason As String, movedByUserId As Integer)
            Using conn = AppStartup.ConnectionFactory.CreateConnection()
                Using tx = conn.BeginTransaction()
                    Try
                        Dim locationId = AppStartup.StorageRepo.GetOrCreateByKey(cab, shelf, box, tx)
                        AppStartup.StorageService.MoveDocument(docId, locationId, movedByUserId, reason, tx)
                        tx.Commit()
                    Catch
                        tx.Rollback()
                        Throw
                    End Try
                End Using
            End Using
        End Sub

        Public Sub RequestRevision(docId As Integer, punchlistNotes As String, returnSection As String, staffName As String, staffUserId As Integer)
            If IsDatabaseConnected AndAlso AppStartup.DirectiveService IsNot Nothing Then
                If TryRunSql("RequestRevision", Sub() AppStartup.DirectiveService.RequestRevision(docId, punchlistNotes, staffUserId, returnSection)) Then
                    PullAfterWrite("Documents", "Directives", "RoutingLogs", "AuditTrail")
                    Return
                End If
            End If
            EmbeddedDB.RequestRevision(docId, punchlistNotes, returnSection, staffName)
        End Sub

        ''' <summary>
        ''' Re-reads the given tables from SQL into the cache on demand. The stale-snapshot
        ''' guard needs server truth in the form's row before the operator retries an
        ''' action the server already applied from another seat.
        ''' </summary>
        Public Sub RefreshFromServer(ParamArray tables As String())
            If Not IsDatabaseConnected Then Return
            PullAfterWrite(tables)
        End Sub

        Public Sub ResubmitDocument(docId As Integer, staffName As String, notes As String, staffUserId As Integer, Optional expectedFromStatusId As Integer = 0)
            If IsDatabaseConnected AndAlso AppStartup.RoutingService IsNot Nothing Then
                If TryRunSql("ResubmitDocument", Sub() AppStartup.RoutingService.ResubmitDocument(docId, staffUserId, notes, expectedFromStatusId)) Then
                    PullAfterWrite("Documents", "RoutingLogs", "Movements", "AuditTrail")
                    Return
                End If
            End If
            EmbeddedDB.ResubmitDocument(docId, staffName, notes)
        End Sub

        Public Sub ApproveDocument(docId As Integer, staffName As String, notes As String, staffUserId As Integer, Optional expectedFromStatusId As Integer = 0)
            If IsDatabaseConnected AndAlso AppStartup.RoutingService IsNot Nothing Then
                If TryRunSql("ApproveDocument", Sub() AppStartup.RoutingService.ApproveDocument(docId, staffUserId, notes, expectedFromStatusId)) Then
                    PullAfterWrite("Documents", "RoutingLogs", "AuditTrail")
                    Return
                End If
            End If
            EmbeddedDB.ApproveDocument(docId, staffName, notes)
        End Sub

        Public Sub ReleaseDocument(docId As Integer, staffName As String, notes As String, staffUserId As Integer, Optional expectedFromStatusId As Integer = 0)
            If IsDatabaseConnected AndAlso AppStartup.RoutingService IsNot Nothing Then
                If TryRunSql("ReleaseDocument", Sub() AppStartup.RoutingService.ReleaseDocument(docId, staffUserId, notes, expectedFromStatusId)) Then
                    PullAfterWrite("Documents", "RoutingLogs", "AuditTrail")
                    Return
                End If
            End If
            EmbeddedDB.ReleaseDocument(docId, staffName, notes)
        End Sub
        ''' <summary>
        ''' Enrols a staff badge. Connected: tbl_Users plus tbl_RfidCards plus tbl_UserRoles,
        ''' so every desk that syncs sees the badge. Offline, or after a failed SQL write: the
        ''' cache, as before. Returns True when the enrolment landed in the shared database.
        ''' </summary>
        Public Function RegisterUser(fullName As String, role As String, office As String, cardUid As String, canRoute As Boolean, canMove As Boolean, canSoftCopy As Boolean, enrolledByName As String, enrolledByUserId As Integer, Optional claimGuard As Boolean = False) As Boolean
            Dim cleanUid = If(cardUid, "").Trim().ToUpperInvariant()
            If IsDatabaseConnected AndAlso AppStartup.UserRepo IsNot Nothing Then
                If EnsureUserInSql(fullName, role, office, cleanUid, canRoute, canMove, canSoftCopy, enrolledByUserId, claimGuard) Then
                    PullAfterWrite("Users", "AuditTrail")
                    Return True
                End If
            End If

            EmbeddedDB.AddUser(cleanUid, fullName, role, office, canRoute, canMove, canSoftCopy, pendingSync:=True)
            EmbeddedDB.LogAudit(enrolledByName, String.Format("Registered/Updated User [{0}] Role: {1} (Desk: {2}) Privileges: [Route:{3}, Move:{4}, SoftCopy:{5}] RFID: ****{6}", fullName, role, office, canRoute, canMove, canSoftCopy, Right(cleanUid, 4)), userId:=enrolledByUserId, actionType:="USER_REGISTERED")
            Return False
        End Function

        ''' <summary>
        ''' Desktop role or section name to SQL role code. The SQL role names are what the
        ''' Users mirror projects back into the cache Role column, so this map is what keeps
        ''' the section desk logic (EmbeddedDB.GetVisibleDocuments) working after a sync.
        ''' </summary>
        Private Shared Function RoleCodeFor(role As String) As String
            Select Case If(role, "").Trim().ToUpperInvariant()
                Case "SECRETARY-GENERAL", RbacPolicy.ROLE_SG : Return RbacPolicy.ROLE_SG
                Case "SYSTEM ADMINISTRATOR", RbacPolicy.ROLE_SYSADMIN : Return RbacPolicy.ROLE_SYSADMIN
                Case "OSG CHIEF" : Return RbacPolicy.ROLE_OSG_CHIEF
                Case "RECORDS SECTION" : Return RbacPolicy.ROLE_RECORDS
                Case "SECRETARIAT" : Return RbacPolicy.ROLE_SECRETARIAT
                Case "LEGISLATIVE SECTION" : Return RbacPolicy.ROLE_LEGISLATIVE
                Case "FINANCE SECTION" : Return RbacPolicy.ROLE_FINANCE
                Case "TRAVEL SECTION" : Return RbacPolicy.ROLE_TRAVEL
                Case Else : Return RbacPolicy.ROLE_ADMIN_STAFF
            End Select
        End Function

        ''' <summary>
        ''' Username has to be stable and unique (tbl_Users.Username is UNIQUE), and the Admin
        ''' screen only asks for a name and a badge, so derive it from the name.
        ''' </summary>
        Private Shared Function SanitizeUsername(fullName As String) As String
            Dim sb As New System.Text.StringBuilder()
            For Each ch In If(fullName, "")
                If Char.IsLetterOrDigit(ch) Then sb.Append(Char.ToLowerInvariant(ch))
            Next
            If sb.Length = 0 Then sb.Append("staff")
            Return sb.ToString(0, Math.Min(sb.Length, 40))
        End Function

        ''' <summary>
        ''' True when the connected action completed. False routes the caller to the cache
        ''' path, so a dead SQL host still lets the desk work instead of losing the action.
        ''' A refused transition (the server-side from-status guard) is NOT a dead host:
        ''' replaying it offline would write phantom custody rows for a move the server
        ''' deliberately rejected, so GuardRefusedException is rethrown to the caller.
        ''' The fallback is ALSO reported to the operator surface: in a WinExe the trace log
        ''' is invisible, and a silent offline fallback looks identical to success.
        ''' </summary>
        Private Function TryRunSql(label As String, action As Action) As Boolean
            Try
                action()
                Return True
            Catch ex As GuardRefusedException
                Throw
            Catch ex As Exception
                ' Drop to the cache immediately: the next action must not pay the same connect
                ' timeout again. The next successful pull restores SQL writes.
                _sqlAvailable = False
                System.Diagnostics.Trace.WriteLine("SQL Server " & label & " error (falling back to embedded cache): " & ex.Message)
                Program.ReportOperatorWarning("SQL write fell back to offline cache (" & label & "): " & ex.Message)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' After a connected write, refresh just the tables that write touched, so the
        ''' operator's own action shows in the grid in the same click. Runs on the caller's
        ''' thread and persists, unlike the background poll.
        ''' </summary>
        Private Sub PullAfterWrite(ParamArray tables As String())
            EmbeddedDB.ApplySnapshots(BuildSqlSnapshot(tables))
        End Sub

    End Class
End Namespace
