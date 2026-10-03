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
    Public Class DesktopDataCoordinator
        ' The connection state at startup only. Writes and outbox replay gate on
        ' IsDatabaseConnected below, so a seat that booted while the server was down recovers
        ' instead of staying on the cache for the rest of the day.
        Private ReadOnly _isDatabaseConnected As Boolean

        Public ReadOnly Property IsDatabaseConnected As Boolean
            Get
                ' The startup probe can report offline while the host was only briefly busy;
                ' once any snapshot pull has confirmed the host live again, the write paths
                ' recover instead of staying stranded offline for the whole session.
                Return _isDatabaseConnected OrElse Program.IsDatabaseConnected
            End Get
        End Property

        Public Sub New(isDatabaseConnected As Boolean)
            _isDatabaseConnected = isDatabaseConnected
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

                Dim doc As Document = Nothing
                ' One connection, one transaction: the sequence reservation, the document row,
                ' the workflow update, the assignment, the routing log, and the storage
                ' movement either land whole or not at all, so a mid-sequence failure cannot
                ' strand an orphan SQL document for the outbox to re-register under a fresh code.
                Using conn = AppStartup.ConnectionFactory.CreateConnection()
                    Using tx = conn.BeginTransaction()
                        Try
                            doc = AppStartup.DocService.RegisterDocumentWithWorkflow(fields.Title, TypeCodeFor(fields.DocType), fields.FlowDirection, fields.Origin, fields.Destination, targetDt, fields.GDriveUrl, fields.LastAction, fields.RegisteredByUserId, transaction:=tx, externalControlNumber:=fields.ExternalControlNumber, preferredDocCode:=fields.PreferredDocCode)
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
                        Catch
                            tx.Rollback()
                            Throw
                        End Try
                    End Using
                End Using

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
                System.Diagnostics.Trace.WriteLine("SQL Server Document Register error (falling back to embedded cache): " & ex.Message)
                Program.ReportOperatorWarning("Registration was saved to this workstation's offline store because the office server refused the write (" & ex.Message & "). It retries automatically; if this repeats, run Station Setup.")
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Scans local cache for records created or updated while offline (PendingSync = True) and
        ''' replays them to SQL Server, minting authoritative sequences, re-mapping child IDs, and
        ''' syncing routing logs, movements, directives, and audit events.
        ''' Must run before snapshot pulls so offline work is never lost.
        ''' </summary>
        Public Function SyncOfflineOutbox() As Integer
            If Not IsDatabaseConnected Then Return 0
            If AppStartup.DocService Is Nothing Then Return 0

            Dim uploadedCount As Integer = 0
            ' Users replay first: offline-enrolled staff carry local autoincrement ids, and
            ' every later replay has to resolve those to the SQL ids before it can attribute
            ' a document, directive, or custody row to the right person.
            SyncLock _replayUserMapLock
                _replayUserMap.Clear()
            End SyncLock
            uploadedCount += ReplayPendingUsers()
            uploadedCount += ReplayNewOfflineDocuments()
            uploadedCount += ReplayOfflineDocumentStateUpdates()

            ' Every replay batch flips its flags in memory only; a flush here closes the gap
            ' where a crash after the SQL commit would replay the batch all over again.
            If uploadedCount > 0 Then EmbeddedDB.Save()

            Dim childCount As Integer = 0
            childCount += ReplayPendingDirectives()
            If childCount > 0 Then EmbeddedDB.Save()
            childCount += ReplayPendingRoutingLogs()
            If childCount > 0 Then EmbeddedDB.Save()
            childCount += ReplayPendingMovements()
            If childCount > 0 Then EmbeddedDB.Save()
            childCount += ReplayPendingAuditTrail()
            If childCount > 0 Then EmbeddedDB.Save()
            Return uploadedCount + childCount
        End Function

        ' Local-to-SQL user id pairs resolved during this sync pass. The cache's ids are
        ' local autoincrement values; writing one into SQL would attribute the row to
        ' whoever happens to hold that id on the server.
        Private ReadOnly _replayUserMap As New Dictionary(Of Integer, Integer)()
        Private ReadOnly _replayUserMapLock As New Object()

        Private Function ResolveReplayUserId(localUserId As Integer) As Integer?
            If localUserId <= 0 Then Return Nothing
            SyncLock _replayUserMapLock
                Dim mapped As Integer = 0
                If _replayUserMap.TryGetValue(localUserId, mapped) Then Return mapped
            End SyncLock

            Dim resolved As Integer? = Nothing
            Dim fullName As String = Nothing
            SyncLock EmbeddedDB.SyncRoot
                Dim dtUsers = EmbeddedDB.DataSet.Tables("Users")
                If dtUsers IsNot Nothing Then
                    Dim rows = dtUsers.Select(String.Format("UserID = {0}", localUserId))
                    If rows.Length > 0 Then fullName = rows(0)("FullName").ToString()
                End If
            End SyncLock
            If fullName IsNot Nothing AndAlso AppStartup.UserRepo IsNot Nothing Then
                For Each u As User In AppStartup.UserRepo.GetAll()
                    If String.Equals(u.FullName, fullName.Trim(), StringComparison.OrdinalIgnoreCase) Then
                        resolved = u.UserID
                        Exit For
                    End If
                Next
            End If
            ' The id may already be a SQL id (rows created on a connected seat whose child
            ' rows replay later), so a direct lookup is the second chance.
            If Not resolved.HasValue AndAlso AppStartup.UserRepo IsNot Nothing Then
                Dim direct = AppStartup.UserRepo.GetById(localUserId)
                If direct IsNot Nothing Then resolved = direct.UserID
            End If

            If resolved.HasValue Then
                SyncLock _replayUserMapLock
                    _replayUserMap(localUserId) = resolved.Value
                End SyncLock
            End If
            Return resolved
        End Function

        ' One operator banner per area per session: replay retries every sync tick, so an
        ' unresolved row must not re-alert the operator every ten seconds.
        Private Shared ReadOnly _warnedReplayAreas As New HashSet(Of String)()

        Private Shared Sub WarnReplayOnce(area As String, message As String)
            System.Diagnostics.Trace.WriteLine("Replay [" & area & "]: " & message)
            SyncLock _warnedReplayAreas
                If _warnedReplayAreas.Contains(area) Then Return
                _warnedReplayAreas.Add(area)
            End SyncLock
            Program.ReportOperatorWarning("Offline sync [" & area & "]: " & message & " The affected records stay queued and retry automatically.")
        End Sub

        Private Function ReplayNewOfflineDocuments() As Integer
            Dim uploadedCount As Integer = 0
            Dim dtDocs = EmbeddedDB.DataSet.Tables("Documents")
            If dtDocs Is Nothing OrElse Not dtDocs.Columns.Contains("PendingSync") Then Return uploadedCount

            Dim isNewCol = dtDocs.Columns.Contains("IsNewOfflineRecord")
            ' Every Select, row read, and row mutation in the Replay* methods runs under the
            ' store lock: the UI thread enumerates and rewrites these same tables under the
            ' same lock (grid binds, snapshot pulls), and an unlocked DataTable write or row
            ' removal racing that enumeration throws. SQL I/O stays OUTSIDE the lock so the
            ' replay never stalls the operator's screen for a network round trip.
            Dim pendingNewDocs As DataRow()
            SyncLock EmbeddedDB.SyncRoot
                pendingNewDocs = dtDocs.Select(If(isNewCol, "PendingSync = True AND IsNewOfflineRecord = True", "PendingSync = True"))
            End SyncLock
            For Each row In pendingNewDocs
                Try
                    Dim oldLocalId As Integer
                    Dim createdBy As Integer?
                    Dim fields As RegistrationInput
                    SyncLock EmbeddedDB.SyncRoot
                        oldLocalId = CInt(row("DocumentID"))
                        createdBy = ReplayUserId(row, "CreatedByUserID")
                        fields = New RegistrationInput With {
                            .DocType = row("DocType").ToString(),
                            .Title = row("Title").ToString(),
                            .Origin = row("OriginatingOffice").ToString(),
                            .Destination = row("DestinationOffice").ToString(),
                            .Cabinet = row("CabinetID").ToString(),
                            .Shelf = row("ShelfNo").ToString(),
                            .Box = row("BoxCode").ToString(),
                            .GDriveUrl = row("GDriveURL").ToString(),
                            .InitialStatus = row("CurrentStatus").ToString(),
                            .AssignedStaff = row("AssignedStaff").ToString(),
                            .FlowDirection = row("FlowDirection").ToString(),
                            .AssignedSection = row("AssignedSection").ToString(),
                            .Deadline = row("TargetDeadlineUTC").ToString(),
                            .LastAction = row("LastActionTaken").ToString(),
                            .ExternalControlNumber = If(row.Table.Columns.Contains("ExternalControlNumber"), row("ExternalControlNumber").ToString(), ""),
                            .PreferredDocCode = row("DocCode").ToString()
                        }
                    End SyncLock

                    ' Registration attributes the row to a real user (RegisteredByUserID is
                    ' a NOT NULL FK); without a resolvable one the row stays queued rather
                    ' than being filed under whoever holds the fallback id.
                    If Not createdBy.HasValue Then
                        WarnReplayOnce("documents", "an offline registration could not be attributed to a known user")
                        Continue For
                    End If
                    fields.RegisteredByUserId = createdBy.Value

                    Dim newDoc = RegisterDocumentInSqlInternal(fields)
                    If newDoc IsNot Nothing Then
                        Dim newSqlDocId = newDoc.DocumentID
                        SyncLock EmbeddedDB.SyncRoot
                            RemapOfflineChildRecords(oldLocalId, newSqlDocId)
                            row("PendingSync") = False
                            If isNewCol Then row("IsNewOfflineRecord") = False
                            uploadedCount += 1
                            Try
                                dtDocs.Rows.Remove(row)
                            Catch
                            End Try
                        End SyncLock
                        ' The sync flags live in the cache file, so a crash between the SQL
                        ' commit and the flush would replay this document as a brand-new
                        ' registration with a fresh DocCode. Persist per row: the document
                        ' replay is the only step whose replay mints new shared records.
                        EmbeddedDB.Save()
                    End If
                Catch ex As Exception
                    System.Diagnostics.Trace.WriteLine("Failed to replay offline new document to SQL: " & ex.Message)
                End Try
            Next
            Return uploadedCount
        End Function

        Private Function ReplayOfflineDocumentStateUpdates() As Integer
            Dim uploadedCount As Integer = 0
            Dim dtDocs = EmbeddedDB.DataSet.Tables("Documents")
            If dtDocs Is Nothing OrElse Not dtDocs.Columns.Contains("PendingSync") Then Return uploadedCount
            If Not dtDocs.Columns.Contains("IsNewOfflineRecord") Then Return uploadedCount

            Dim pendingUpdatedDocs As DataRow()
            SyncLock EmbeddedDB.SyncRoot
                pendingUpdatedDocs = dtDocs.Select("PendingSync = True AND (IsNewOfflineRecord = False OR IsNewOfflineRecord IS NULL)")
            End SyncLock
            For Each row In pendingUpdatedDocs
                Try
                    Dim docId As Integer
                    Dim statusStr As String
                    SyncLock EmbeddedDB.SyncRoot
                        docId = CInt(row("DocumentID"))
                        statusStr = row("CurrentStatus").ToString()
                    End SyncLock
                    Dim st = AppStartup.ReferenceDataRepo.GetStatusByCode(statusStr)
                    Dim statusId As Integer
                    If st IsNot Nothing Then
                        statusId = st.StatusID
                    Else
                        ' The mirror status is not a known code (legacy free text like
                        ' "SG Directive: ..."): keep the SQL status and replay only the
                        ' section, punchlist, and last-action fields instead of resetting
                        ' the document to RECEIVED.
                        Dim currentDoc = AppStartup.DocumentRepo.GetById(docId)
                        If currentDoc Is Nothing Then
                            SyncLock EmbeddedDB.SyncRoot
                                row("PendingSync") = False
                            End SyncLock
                            Continue For
                        End If
                        statusId = currentDoc.StatusID
                    End If
                    Dim assignedSec As String
                    Dim punchlist As String
                    Dim lastAction As String
                    Dim destOffice As String
                    Dim originOffice As String
                    Dim modifiedBy As Integer?
                    Dim rowVersion As Byte() = Nothing
                    SyncLock EmbeddedDB.SyncRoot
                        assignedSec = row("AssignedSection").ToString()
                        punchlist = row("RevisionPunchlist").ToString()
                        lastAction = row("LastActionTaken").ToString()
                        destOffice = row("DestinationOffice").ToString()
                        originOffice = row("OriginatingOffice").ToString()
                        modifiedBy = ReplayUserId(row, "ModifiedByUserID")
                        If Not modifiedBy.HasValue OrElse modifiedBy.Value = 1 Then modifiedBy = ReplayUserId(row, "CreatedByUserID")
                        If dtDocs.Columns.Contains("RowVersion") AndAlso Not IsDBNull(row("RowVersion")) Then
                            rowVersion = CType(row("RowVersion"), Byte())
                        End If
                    End SyncLock

                    Dim applied As Integer
                    If rowVersion IsNot Nothing Then
                        ' Optimistic guard: the update lands only if the server row is still
                        ' the version this seat mirrored before it went offline.
                        applied = AppStartup.DocumentRepo.UpdateWorkflowStateChecked(docId, rowVersion, statusId, assignedSec, punchlist, lastAction, If(modifiedBy.HasValue, modifiedBy.Value, 0), destinationOffice:=destOffice, originOffice:=originOffice)
                    Else
                        ' Cache predates the RowVersion mirror: replay without the guard, as before.
                        AppStartup.DocumentRepo.UpdateWorkflowState(docId, statusId, assignedSec, punchlist, lastAction, If(modifiedBy.HasValue, modifiedBy.Value, 0), destinationOffice:=destOffice, originOffice:=originOffice)
                        applied = 1
                    End If

                    SyncLock EmbeddedDB.SyncRoot
                        row("PendingSync") = False
                    End SyncLock
                    If applied = 0 Then
                        ' Conflict: another seat moved the document past the mirrored version
                        ' while this one was offline. Server truth stands and the next pull
                        ' restamps the local row; the offline edit is recorded, not applied.
                        Dim conflictCode As String
                        SyncLock EmbeddedDB.SyncRoot
                            conflictCode = row("DocCode").ToString()
                        End SyncLock
                        EmbeddedDB.LogAudit("SYSTEM", "Replay conflict on " & conflictCode & ": server kept the newer copy, offline edit discarded (" & Environment.MachineName & ")", actionType:="SYNC_CONFLICT")
                        Program.ReportOperatorWarning("Sync conflict on " & conflictCode & ": the server had a newer update, so this seat's offline edit was discarded.")
                    Else
                        uploadedCount += 1
                    End If
                Catch ex As Exception
                    System.Diagnostics.Trace.WriteLine("Failed to replay offline document state update to SQL: " & ex.Message)
                End Try
            Next
            Return uploadedCount
        End Function

        ' Offline user enrolment and edits replay here: without this the badge worked only
        ' on the enrolling workstation and the row was wiped by the next Users pull.
        Private Function ReplayPendingUsers() As Integer
            Dim uploadedCount As Integer = 0
            Dim dtUsers = EmbeddedDB.DataSet.Tables("Users")
            If dtUsers Is Nothing OrElse Not dtUsers.Columns.Contains("PendingSync") Then Return uploadedCount

            Dim pendingUsers As DataRow()
            SyncLock EmbeddedDB.SyncRoot
                pendingUsers = dtUsers.Select("PendingSync = True")
            End SyncLock
            For Each row In pendingUsers
                Try
                    Dim fullName As String
                    Dim role As String
                    Dim office As String
                    Dim cardUid As String
                    Dim canRoute As Boolean
                    Dim canMove As Boolean
                    Dim canSoftCopy As Boolean
                    Dim enrolledBy As Integer?
                    SyncLock EmbeddedDB.SyncRoot
                        fullName = row("FullName").ToString()
                        role = row("Role").ToString()
                        office = If(row.Table.Columns.Contains("Office") AndAlso Not IsDBNull(row("Office")), row("Office").ToString(), "")
                        ' The real card number rides in RFID_UID_RAW while the row is pending
                        ' (the hash is all a synced row keeps); pending rows from caches
                        ' written before that column still carry the raw value in RFID_UID.
                        cardUid = If(row.Table.Columns.Contains("RFID_UID_RAW") AndAlso Not IsDBNull(row("RFID_UID_RAW")) AndAlso row("RFID_UID_RAW").ToString().Length > 0,
                                     row("RFID_UID_RAW").ToString(),
                                     row("RFID_UID").ToString())
                        canRoute = If(row.Table.Columns.Contains("CanRoute") AndAlso Not IsDBNull(row("CanRoute")), CBool(row("CanRoute")), True)
                        canMove = If(row.Table.Columns.Contains("CanMove") AndAlso Not IsDBNull(row("CanMove")), CBool(row("CanMove")), True)
                        canSoftCopy = If(row.Table.Columns.Contains("CanSoftCopy") AndAlso Not IsDBNull(row("CanSoftCopy")), CBool(row("CanSoftCopy")), True)
                        enrolledBy = ReplayUserId(row, "UserID")
                    End SyncLock
                    ' The enroller falls back to 0, the shipped "no enroller" convention the
                    ' claim flow already passes: this row IS the new user, so its local id
                    ' cannot resolve in SQL yet and failing it closed would strand the
                    ' enrolment forever. Action rows (documents, custody, directives) keep
                    ' the strict resolution; enrolment does not.
                    Dim enrolledById As Integer = If(enrolledBy.HasValue, enrolledBy.Value, 0)
                    Dim ok As Boolean = EnsureUserInSql(fullName, role, office, cardUid, canRoute, canMove, canSoftCopy, enrolledById)
                    If ok Then
                        SyncLock EmbeddedDB.SyncRoot
                            row("PendingSync") = False
                            If row.Table.Columns.Contains("RFID_UID_RAW") Then row("RFID_UID_RAW") = Nothing
                        End SyncLock
                        uploadedCount += 1
                    End If
                Catch ex As Exception
                    WarnReplayOnce("users", "offline user replay failed: " & ex.Message)
                End Try
            Next
            Return uploadedCount
        End Function

        ''' <summary>
        ''' The connected half of user enrolment: upsert into tbl_Users, issue the badge,
        ''' and assign the role, all in one transaction so a mid-way failure cannot leave
        ''' a user with no card or no role. Shared by RegisterUser and ReplayPendingUsers.
        ''' Returns False so the caller keeps the row pending instead of losing it.
        ''' </summary>
        Private Function EnsureUserInSql(fullName As String, role As String, office As String, cardUid As String, canRoute As Boolean, canMove As Boolean, canSoftCopy As Boolean, enrolledByUserId As Integer) As Boolean
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
                            End If
                            tx.Commit()
                            Return True
                        Catch
                            tx.Rollback()
                            Throw
                        End Try
                    End Using
                End Using
            Catch ex As Exception
                System.Diagnostics.Trace.WriteLine("SQL Server user enrolment error: " & ex.Message)
                Return False
            End Try
        End Function

        Private Function ReplayPendingDirectives() As Integer
            Dim uploadedCount As Integer = 0
            Dim dtDirectives = EmbeddedDB.DataSet.Tables("Directives")
            If dtDirectives Is Nothing OrElse Not dtDirectives.Columns.Contains("PendingSync") Then Return uploadedCount

            Dim pendingDirectives As DataRow()
            SyncLock EmbeddedDB.SyncRoot
                pendingDirectives = dtDirectives.Select("PendingSync = True")
            End SyncLock
            For Each row In pendingDirectives
                Try
                    Dim docId As Integer
                    Dim directiveText As String
                    Dim notes As String
                    Dim issuedBy As Integer?
                    Dim code As String
                    SyncLock EmbeddedDB.SyncRoot
                        docId = CInt(row("DocumentID"))
                        directiveText = row("SGDirective").ToString()
                        notes = row("Notes").ToString()
                        issuedBy = ReplayUserId(row, "IssuedByUserID")
                        code = If(row.Table.Columns.Contains("DirectiveCode") AndAlso Not IsDBNull(row("DirectiveCode")), row("DirectiveCode").ToString(), "")
                    End SyncLock
                    ' The parent document must already exist in SQL: inserting against the
                    ' local id would attach the directive to whatever document holds that id
                    ' on the server. Rows stay queued until their document replays.
                    If AppStartup.DocumentRepo.GetById(docId) Is Nothing Then Continue For
                    If Not issuedBy.HasValue Then
                        WarnReplayOnce("directives", "an offline directive could not be attributed to a known user")
                        Continue For
                    End If

                    ' Map the stored directive code to the seeded type; the doc state itself
                    ' is replayed by ReplayOfflineDocumentStateUpdates, so this insert must
                    ' not re-apply a status side effect (IssueDirective would overwrite the
                    ' replayed state with the type's ResultStatusID).
                    Dim directiveTypeId As Integer = 1
                    If code.Length > 0 Then
                        For Each dirType As DirectiveType In AppStartup.ReferenceDataRepo.GetDirectiveTypes()
                            If String.Equals(dirType.DirectiveCode, code, StringComparison.OrdinalIgnoreCase) Then
                                directiveTypeId = dirType.DirectiveTypeID
                                Exit For
                            End If
                        Next
                    End If

                    Dim directive As New ActionDirective With {
                        .DocumentID = docId,
                        .DirectiveTypeID = directiveTypeId,
                        .DirectiveText = directiveText,
                        .IssuedByUserID = issuedBy.Value,
                        .IssuedAtUTC = DateTime.UtcNow,
                        .IsActive = True,
                        .Remarks = notes
                    }
                    AppStartup.DirectiveRepo.Insert(directive)
                    If AppStartup.AuditService IsNot Nothing Then
                        AppStartup.AuditService.LogEvent("DIRECTIVE_ADDED", "Directive", docId.ToString(), Nothing, Nothing, Nothing, True, Nothing)
                    End If
                    SyncLock EmbeddedDB.SyncRoot
                        row("PendingSync") = False
                    End SyncLock
                    uploadedCount += 1
                Catch ex As Exception
                    WarnReplayOnce("directives", "offline directive replay failed: " & ex.Message)
                End Try
            Next
            Return uploadedCount
        End Function

        Private Function ReplayPendingRoutingLogs() As Integer
            Dim uploadedCount As Integer = 0
            Dim dtRouting = EmbeddedDB.DataSet.Tables("RoutingLogs")
            If dtRouting Is Nothing OrElse Not dtRouting.Columns.Contains("PendingSync") Then Return uploadedCount

            Dim pendingRouting As DataRow()
            SyncLock EmbeddedDB.SyncRoot
                pendingRouting = dtRouting.Select("PendingSync = True")
            End SyncLock
            For Each row In pendingRouting
                Try
                    Dim action As String
                    SyncLock EmbeddedDB.SyncRoot
                        action = row("ActionTaken").ToString()
                    End SyncLock
                    If Not action.Equals("REGISTERED", StringComparison.OrdinalIgnoreCase) Then
                        Dim docId As Integer
                        Dim fromOffice As String
                        Dim toOffice As String
                        Dim remarks As String
                        Dim routedBy As Integer?
                        SyncLock EmbeddedDB.SyncRoot
                            docId = CInt(row("DocumentID"))
                            fromOffice = row("FromOffice").ToString()
                            toOffice = row("ToOffice").ToString()
                            remarks = row("Remarks").ToString()
                            routedBy = ReplayUserId(row, "RoutedByUserID")
                        End SyncLock
                        If AppStartup.DocumentRepo.GetById(docId) Is Nothing Then Continue For
                        If Not routedBy.HasValue Then
                            WarnReplayOnce("routing", "an offline routing log could not be attributed to a known user")
                            Continue For
                        End If
                        ' Only a real status code becomes ToStatusID; mapping free-text
                        ' actions onto RECEIVED would corrupt the custody trail's status,
                        ' so unknown actions get a NULL status instead.
                        Dim st = AppStartup.ReferenceDataRepo.GetStatusByCode(action)
                        Dim statusId As Integer? = If(st IsNot Nothing, CType(st.StatusID, Integer?), Nothing)

                        Dim logEntry As New RoutingLog With {
                            .DocumentID = docId,
                            .FromStatusID = Nothing,
                            .ToStatusID = statusId,
                            .FromOffice = fromOffice,
                            .ToOffice = toOffice,
                            .RoutingRemarks = remarks,
                            .RoutedByUserID = routedBy.Value,
                            .RoutedAtUTC = DateTime.UtcNow
                        }
                        AppStartup.RoutingRepo.Insert(logEntry)
                    End If
                    SyncLock EmbeddedDB.SyncRoot
                        row("PendingSync") = False
                    End SyncLock
                    uploadedCount += 1
                Catch ex As Exception
                    WarnReplayOnce("routing", "offline routing log replay failed: " & ex.Message)
                End Try
            Next
            Return uploadedCount
        End Function

        Private Function ReplayPendingMovements() As Integer
            Dim uploadedCount As Integer = 0
            Dim dtMovements = EmbeddedDB.DataSet.Tables("Movements")
            If dtMovements Is Nothing OrElse Not dtMovements.Columns.Contains("PendingSync") Then Return uploadedCount

            Dim pendingMovements As DataRow()
            SyncLock EmbeddedDB.SyncRoot
                pendingMovements = dtMovements.Select("PendingSync = True")
            End SyncLock
            For Each row In pendingMovements
                Try
                    Dim reason As String
                    SyncLock EmbeddedDB.SyncRoot
                        reason = row("Reason").ToString()
                    End SyncLock
                    If Not reason.Equals("Initial storage assignment.", StringComparison.OrdinalIgnoreCase) Then
                        Dim docId As Integer
                        Dim toLoc As String
                        Dim movedBy As Integer?
                        SyncLock EmbeddedDB.SyncRoot
                            docId = CInt(row("DocumentID"))
                            toLoc = row("ToLocation").ToString()
                            movedBy = ReplayUserId(row, "MovedByUserID")
                        End SyncLock
                        If AppStartup.DocumentRepo.GetById(docId) Is Nothing Then Continue For
                        If Not movedBy.HasValue Then
                            WarnReplayOnce("movements", "an offline storage movement could not be attributed to a known user")
                            Continue For
                        End If
                        Dim parts = toLoc.Split(New Char() {"/"c, "|"c})
                        Dim cab = If(parts.Length > 0, parts(0).Trim(), "")
                        Dim shelf = If(parts.Length > 1, parts(1).Trim(), "")
                        Dim box = If(parts.Length > 2, parts(2).Trim(), "")

                        Dim locationId = AppStartup.StorageRepo.GetOrCreateByKey(cab, shelf, box)

                        AppStartup.StorageService.MoveDocument(docId, locationId, movedBy.Value, reason)
                    End If
                    SyncLock EmbeddedDB.SyncRoot
                        row("PendingSync") = False
                    End SyncLock
                    uploadedCount += 1
                Catch ex As Exception
                    WarnReplayOnce("movements", "offline movement replay failed: " & ex.Message)
                End Try
            Next
            Return uploadedCount
        End Function

        Private Function ReplayPendingAuditTrail() As Integer
            Dim uploadedCount As Integer = 0
            Dim dtAudit = EmbeddedDB.DataSet.Tables("AuditTrail")
            If dtAudit Is Nothing OrElse Not dtAudit.Columns.Contains("PendingSync") Then Return uploadedCount

            Dim pendingAudit As DataRow()
            SyncLock EmbeddedDB.SyncRoot
                pendingAudit = dtAudit.Select("PendingSync = True")
            End SyncLock
            For Each row In pendingAudit
                Try
                    Dim actionDesc As String
                    Dim userId As Integer?
                    Dim originalType As String
                    Dim actor As String
                    Dim auditId As String
                    SyncLock EmbeddedDB.SyncRoot
                        actionDesc = row("ActionDescription").ToString()
                        userId = ReplayUserId(row, "UserID")
                        ' Keep the original action type and actor so the SQL audit trail reads
                        ' as the events that happened, not as one opaque OFFLINE_SYNC blob.
                        originalType = If(row.Table.Columns.Contains("ActionType") AndAlso Not IsDBNull(row("ActionType")), row("ActionType").ToString(), "")
                        actor = If(row.Table.Columns.Contains("UserName") AndAlso Not IsDBNull(row("UserName")), row("UserName").ToString(), "SYSTEM")
                        auditId = row("AuditID").ToString()
                    End SyncLock
                    AppStartup.AuditRepo.Insert(New AuditEntry With {
                        .UserID = userId,
                        .UsernameSnapshot = actor,
                        .ActionType = If(originalType.Length > 0, originalType, "OFFLINE_SYNC"),
                        .EntityType = If(originalType.Length > 0, "OFFLINE_REPLAY", "SYNC_REPLAY"),
                        .EntityID = auditId,
                        .NewValuesJson = actionDesc,
                        .Success = True,
                        .MachineName = Environment.MachineName,
                        .EventAtUTC = DateTime.UtcNow
                    })
                    SyncLock EmbeddedDB.SyncRoot
                        row("PendingSync") = False
                    End SyncLock
                    uploadedCount += 1
                Catch ex As Exception
                    WarnReplayOnce("audit", "offline audit replay failed: " & ex.Message)
                End Try
            Next
            Return uploadedCount
        End Function

        ''' <summary>
        ''' Reads a cached user-id column and resolves it to the SQL user id. The cache's
        ''' ids are local autoincrement values (or the sentinel 1 from very old rows), so
        ''' the value is never written through as-is: callers that attribute a NOT NULL FK
        ''' column leave the row queued when this returns Nothing instead of filing the
        ''' record under an arbitrary account.
        ''' </summary>
        Private Function ReplayUserId(row As DataRow, columnName As String) As Integer?
            If row.Table.Columns.Contains(columnName) AndAlso Not IsDBNull(row(columnName)) Then
                Return ResolveReplayUserId(CInt(row(columnName)))
            End If
            Return Nothing
        End Function

        Friend Sub RemapOfflineChildRecords(oldDocId As Integer, newDocId As Integer)
            For Each tblName In {"Directives", "RoutingLogs", "Movements"}
                Dim table = EmbeddedDB.DataSet.Tables(tblName)
                If table IsNot Nothing AndAlso table.Columns.Contains("DocumentID") Then
                    For Each r In table.Select(String.Format("DocumentID = {0}", oldDocId))
                        r("DocumentID") = newDocId
                    Next
                End If
            Next
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
            Public Property LastAction As String
            Public Property RegisteredByUserId As Integer
            Public Property ExternalControlNumber As String
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
        Private Shared Function ResolveUserId(fullName As String) As Integer?
            If String.IsNullOrWhiteSpace(fullName) OrElse AppStartup.UserRepo Is Nothing Then Return Nothing
            For Each u As User In AppStartup.UserRepo.GetAll()
                If String.Equals(u.FullName, fullName.Trim(), StringComparison.OrdinalIgnoreCase) Then Return u.UserID
            Next
            Return Nothing
        End Function

        Public Sub RouteDocument(docId As Integer, fromOffice As String, toOffice As String, action As String, remarks As String, routedByName As String, routedByUserId As Integer)
            If IsDatabaseConnected AndAlso AppStartup.RoutingService IsNot Nothing Then
                ' Status update and custody log are one fact: a crash between them must not
                ' leave a moved document with no routing entry, so both join one transaction.
                If TryRunSql("RouteDocument", Sub()
                                                  Using conn = AppStartup.ConnectionFactory.CreateConnection()
                                                      Using tx = conn.BeginTransaction()
                                                          Try
                                                              Dim status = AppStartup.ReferenceDataRepo.GetStatusByCode(If(String.IsNullOrWhiteSpace(action), "ROUTED", action))
                                                              ' An action that is not a status code must not regress the
                                                              ' document: keep its current status, exactly as the offline
                                                              ' and replay paths do. Only the destination and custody log
                                                              ' record the transmittal.
                                                              Dim statusId As Integer
                                                              If status IsNot Nothing Then
                                                                  statusId = status.StatusID
                                                              Else
                                                                  Dim currentDoc = AppStartup.DocumentRepo.GetById(docId)
                                                                  If currentDoc Is Nothing Then Throw New InvalidOperationException("Document " & docId.ToString() & " no longer exists.")
                                                                  statusId = currentDoc.StatusID
                                                              End If

                                                              AppStartup.DocumentRepo.UpdateWorkflowState(docId, statusId, toOffice, Nothing, "Routed to " & toOffice & ": " & action, routedByUserId, destinationOffice:=toOffice, originOffice:=fromOffice, transaction:=tx)

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
                                              End Sub) Then
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
                If TryRunSql("ApplyDirective", Sub() AppStartup.DirectiveService.IssueDirective(docId, directiveTypeId, directiveText, notes, staffUserId)) Then
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
                If TryRunSql("MoveStorage", Sub()
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
                                            End Sub) Then
                    PullAfterWrite("Documents", "Movements", "AuditTrail")
                    Return
                End If
            End If
            EmbeddedDB.AddMovementLog(docId, fromLoc, toLoc, movedByName, reason)
            EmbeddedDB.LogAudit(movedByName, String.Format("Transferred Doc #{0} storage location from {1} to {2}", docId, fromLoc, toLoc))
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
        Public Function RegisterUser(fullName As String, role As String, office As String, cardUid As String, canRoute As Boolean, canMove As Boolean, canSoftCopy As Boolean, enrolledByName As String, enrolledByUserId As Integer) As Boolean
            Dim cleanUid = If(cardUid, "").Trim().ToUpperInvariant()
            If IsDatabaseConnected AndAlso AppStartup.UserRepo IsNot Nothing Then
                If EnsureUserInSql(fullName, role, office, cleanUid, canRoute, canMove, canSoftCopy, enrolledByUserId) Then
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
        ''' The fallback is ALSO reported to the operator surface: in a WinExe the trace log
        ''' is invisible, and a silent offline fallback looks identical to success.
        ''' </summary>
        Private Function TryRunSql(label As String, action As Action) As Boolean
            Try
                action()
                Return True
            Catch ex As Exception
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

        ''' <summary>
        ''' Reads the shared SQL Server state into cache-shaped tables, one per requested
        ''' cache table name. Runs off the UI thread: it only reads and clones, it never
        ''' touches EmbeddedDB.DataSet rows. Returns Nothing when the host is known to be
        ''' offline or the connection fails, so the poller can skip a tick silently.
        ''' A projection that fails is traced and skipped rather than taking the other five
        ''' down with it: the mirror's failure mode is silent, so one stale table beats a
        ''' mirror that stopped working with no evidence anywhere.
        ''' </summary>
        Public Function BuildSqlSnapshot(ParamArray tables As String()) As Dictionary(Of String, DataTable)
            If AppStartup.Settings Is Nothing OrElse Not AppStartup.Settings.DatabaseSettings.UseSqlServer Then Return Nothing

            ' Every table is pulled in full on each tick. The merge ignores unchanged rows, so
            ' the UI cost of a quiet tick is one indexed lookup per row. That stays fine to
            ' thousands of documents; when AuditTrail reaches six figures, switch this to an
            ' incremental pull keyed on AuditID with an append-only merge so the per-tick cost
            ' tracks new rows instead of total rows.
            Dim names = If(tables Is Nothing OrElse tables.Length = 0,
                           New String() {"Documents", "Users", "Directives", "RoutingLogs", "Movements", "AuditTrail", "Heartbeat"},
                           tables)
            Dim result As New Dictionary(Of String, DataTable)()
            Try
                ' Connect with a 5s budget so an unreachable host skips a tick promptly. The
                ' attempt doubles as the connection re-probe: the startup probe runs once, so
                ' without this a host that was busy at launch would stay "offline" all day.
                Dim builder As New Microsoft.Data.SqlClient.SqlConnectionStringBuilder(AppStartup.Settings.DatabaseSettings.ConnectionString) With {
                    .ConnectTimeout = 5
                }
                Using conn As New Microsoft.Data.SqlClient.SqlConnection(builder.ConnectionString)
                    conn.Open()
                    Program.IsDatabaseConnected = True
                    Try
                        UpsertWorkstationHeartbeat(conn)
                    Catch ex As Exception
                        ' A server provisioned before the heartbeat table existed must still
                        ' sync; the heartbeat is decoration, so an older schema only costs a
                        ' trace line.
                        System.Diagnostics.Trace.WriteLine("Workstation heartbeat skipped: " & ex.Message)
                    End Try
                    For Each name In names
                        Dim sql = SqlFor(name)
                        If sql Is Nothing Then Continue For
                        Try
                            Using cmd As New Microsoft.Data.SqlClient.SqlCommand(sql, conn)
                                Using reader = cmd.ExecuteReader()
                                    Dim table As DataTable
                                    SyncLock EmbeddedDB.SyncRoot
                                        table = EmbeddedDB.DataSet.Tables(name).Clone()
                                    End SyncLock
                                    table.Load(reader)
                                    If table.Rows.Count > 0 Then
                                        result(name) = table
                                    End If
                                End Using
                            End Using
                        Catch ex As Exception
                            System.Diagnostics.Trace.WriteLine("SQL snapshot read failed for " & name & " (table left as cached): " & ex.Message)
                        End Try
                    Next
                End Using
            Catch ex As Exception
                Program.IsDatabaseConnected = False
                System.Diagnostics.Trace.WriteLine("SQL snapshot connection failed: " & ex.Message)
                Return Nothing
            End Try
            Return result
        End Function

        ''' <summary>
        ''' One row per seat in tbl_WorkstationHeartbeat, refreshed on every snapshot pull so
        ''' the Admin tab's Workstation Sync Status grid shows who is alive, on which app
        ''' version, and who is accumulating offline work. Called on an open connection;
        ''' callers own the error handling so an older server schema degrades quietly.
        ''' </summary>
        Private Shared Sub UpsertWorkstationHeartbeat(conn As Microsoft.Data.SqlClient.SqlConnection)
            Dim version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version
            Using cmd As New Microsoft.Data.SqlClient.SqlCommand(
                "IF EXISTS (SELECT 1 FROM dbo.tbl_WorkstationHeartbeat WHERE MachineName = @m) " &
                "UPDATE dbo.tbl_WorkstationHeartbeat SET LastSyncUTC = SYSUTCDATETIME(), AppVersion = @v, PendingOutbox = @p, LastError = @e WHERE MachineName = @m " &
                "ELSE INSERT INTO dbo.tbl_WorkstationHeartbeat (MachineName, LastSyncUTC, AppVersion, PendingOutbox, LastError) " &
                "VALUES (@m, SYSUTCDATETIME(), @v, @p, @e)", conn)
                cmd.Parameters.AddWithValue("@m", Environment.MachineName)
                cmd.Parameters.AddWithValue("@v", If(version IsNot Nothing, version.ToString(), ""))
                cmd.Parameters.AddWithValue("@p", PendingOutboxCount())
                cmd.Parameters.AddWithValue("@e", If(EmbeddedDB.LastPersistenceError, ""))
                cmd.ExecuteNonQuery()
            End Using
        End Sub

        Private Shared Function PendingOutboxCount() As Integer
            Dim total As Integer = 0
            SyncLock EmbeddedDB.SyncRoot
                For Each name In {"Documents", "Users", "Directives", "RoutingLogs", "Movements", "AuditTrail"}
                    Dim table = EmbeddedDB.DataSet.Tables(name)
                    If table IsNot Nothing AndAlso table.Columns.Contains("PendingSync") Then
                        total += table.Select("PendingSync = True").Length
                    End If
                Next
            End SyncLock
            Return total
        End Function

        ''' <summary>
        ''' One projection per cache table, shaped to the columns EmbeddedDB.CreateTables
        ''' defines. Nothing is parameterized because nothing varies; every value the UI
        ''' filters on (StatusCode, TypeName, section, full name) is mirrored verbatim so
        ''' the DataView filters and ComboBox lookups keep working untouched. An unknown
        ''' table name returns Nothing, and the caller skips it.
        ''' </summary>
        Private Shared Function SqlFor(tableName As String) As String
            Select Case tableName
                Case "Documents"
                    Return "SELECT d.DocumentID, d.DocCode, ISNULL(dt.TypeName, 'Regular Communication') AS DocType, d.Title, " &
                           "ISNULL(d.OriginOffice, '') AS OriginatingOffice, ISNULL(d.DestinationOffice, '') AS DestinationOffice, " &
                           "ISNULL(sl.CabinetID, '') AS CabinetID, ISNULL(sl.ShelfNo, '') AS ShelfNo, ISNULL(sl.BoxCode, '') AS BoxCode, " &
                           "ISNULL(d.GoogleDriveUrl, '') AS GDriveURL, ISNULL(st.StatusCode, 'RECEIVED') AS CurrentStatus, " &
                           "ISNULL(a.FullName, '') AS AssignedStaff, " &
                           "CONVERT(varchar(19), ISNULL(d.ReceivedDate, CAST(d.RegisteredAtUTC AS date)), 120) AS DateReceived, " &
                           "ISNULL(d.FlowDirection, 'INCOMING') AS FlowDirection, ISNULL(d.AssignedSection, '') AS AssignedSection, " &
                           "CONVERT(varchar(19), d.TargetDeadlineUTC, 120) AS TargetDeadlineUTC, " &
                           "ISNULL(d.RevisionPunchlist, '') AS RevisionPunchlist, ISNULL(d.LastActionTaken, '') AS LastActionTaken, " &
                           "ISNULL(d.ExternalControlNumber, '') AS ExternalControlNumber, d.RowVersion " &
                           "FROM dbo.tbl_Documents d " &
                           "LEFT JOIN dbo.tbl_DocumentTypes dt ON dt.DocumentTypeID = d.DocumentTypeID " &
                           "LEFT JOIN dbo.tbl_DocumentStatuses st ON st.StatusID = d.StatusID " &
                           "LEFT JOIN dbo.tbl_StorageLocations sl ON sl.StorageLocationID = d.CurrentStorageLocationID " &
                           "OUTER APPLY (SELECT TOP 1 u.FullName FROM dbo.tbl_DocumentAssignments da " &
                           "  JOIN dbo.tbl_Users u ON u.UserID = da.AssignedUserID " &
                           "  WHERE da.DocumentID = d.DocumentID ORDER BY da.AssignmentID DESC) a " &
                           "WHERE d.IsDeleted = 0 ORDER BY d.DocumentID"
                Case "Heartbeat"
                    Return "SELECT MachineName, CONVERT(varchar(19), LastSyncUTC, 120) AS LastSyncUTC, " &
                           "ISNULL(AppVersion, '') AS AppVersion, PendingOutbox, ISNULL(LastError, '') AS LastError " &
                           "FROM dbo.tbl_WorkstationHeartbeat ORDER BY MachineName"
                Case "Users"
                    ' RFID_UID is masked for display, RFID_UID_HASH is what authentication
                    ' matches (the same SHA-256 digest EmbeddedDB.Sha256Hex computes over the
                    ' uppercase ASCII card id), and RFID_UID_RAW is cleared: only a row still
                    ' pending offline replay keeps the real card number on disk.
                    Return "SELECT u.UserID, CASE WHEN c.CardPublicID IS NULL THEN '' ELSE '****' + RIGHT(c.CardPublicID, 4) END AS RFID_UID, " &
                           "CONVERT(varchar(64), HASHBYTES('SHA2_256', CAST(ISNULL(c.CardPublicID, '') AS varchar(50))), 2) AS RFID_UID_HASH, " &
                           "CAST(NULL AS varchar(50)) AS RFID_UID_RAW, ISNULL(u.FullName, '') AS FullName, " &
                           "ISNULL(ro.RoleName, ISNULL(u.Office, 'Administrative Staff')) AS Role, ISNULL(u.Office, '') AS Office, " &
                           "ISNULL(u.CanRoute, 1) AS CanRoute, ISNULL(u.CanMove, 1) AS CanMove, ISNULL(u.CanSoftCopy, 1) AS CanSoftCopy, ISNULL(u.IsActive, 1) AS IsActive, " &
                           "ISNULL(u.IsLocked, 0) AS IsLocked " &
                           "FROM dbo.tbl_Users u " &
                           "OUTER APPLY (SELECT TOP 1 rc.CardPublicID FROM dbo.tbl_RfidCards rc WHERE rc.UserID = u.UserID " &
                           "  AND rc.IsActive = 1 AND rc.RevokedAtUTC IS NULL ORDER BY rc.RfidCardID DESC) c " &
                           "OUTER APPLY (SELECT TOP 1 rr.RoleName FROM dbo.tbl_UserRoles ur " &
                           "  JOIN dbo.tbl_Roles rr ON rr.RoleID = ur.RoleID " &
                           "  WHERE ur.UserID = u.UserID AND ur.IsActive = 1 AND rr.IsActive = 1 ORDER BY ur.UserRoleID DESC) ro " &
                           "WHERE u.IsActive = 1 ORDER BY u.UserID"
                Case "Directives"
                    Return "SELECT ad.DirectiveID, ad.DocumentID, ISNULL(ad.DirectiveText, '') AS SGDirective, " &
                           "ISNULL(u.FullName, '') AS AssignedTo, ISNULL(ad.Remarks, '') AS Notes, " &
                           "ISNULL(u.FullName, 'SYSTEM') AS LogUser, CONVERT(varchar(19), ad.IssuedAtUTC, 120) AS Timestamp " &
                           "FROM dbo.tbl_ActionDirectives ad LEFT JOIN dbo.tbl_Users u ON u.UserID = ad.IssuedByUserID " &
                           "ORDER BY ad.DirectiveID"
                Case "RoutingLogs"
                    Return "SELECT r.RoutingLogID AS RoutingID, r.DocumentID, ISNULL(r.FromOffice, '') AS FromOffice, " &
                           "ISNULL(r.ToOffice, '') AS ToOffice, ISNULL(u.FullName, 'SYSTEM') AS RoutedBy, " &
                           "ISNULL(st.StatusCode, '') AS ActionTaken, ISNULL(r.RoutingRemarks, '') AS Remarks, " &
                           "CONVERT(varchar(19), r.RoutedAtUTC, 120) AS Timestamp " &
                           "FROM dbo.tbl_RoutingLogs r LEFT JOIN dbo.tbl_Users u ON u.UserID = r.RoutedByUserID " &
                           "LEFT JOIN dbo.tbl_DocumentStatuses st ON st.StatusID = r.ToStatusID " &
                           "ORDER BY r.RoutingLogID"
                Case "Movements"
                    Return "SELECT m.MovementID, m.DocumentID, '' AS FromLocation, " &
                           "ISNULL(sl.CabinetID, '') + '|' + ISNULL(sl.ShelfNo, '') + '|' + ISNULL(sl.BoxCode, '') AS ToLocation, " &
                           "ISNULL(u.FullName, 'SYSTEM') AS MovedBy, ISNULL(m.MovementReason, '') AS Reason, " &
                           "CONVERT(varchar(19), m.MovedAtUTC, 120) AS Timestamp " &
                           "FROM dbo.tbl_DocumentMovements m LEFT JOIN dbo.tbl_Users u ON u.UserID = m.MovedByUserID " &
                           "LEFT JOIN dbo.tbl_StorageLocations sl ON sl.StorageLocationID = m.StorageLocationID " &
                           "ORDER BY m.MovementID"
                Case "AuditTrail"
                    ' The audit trail is append-only; mirror the newest 1000 events in chronological order.
                    Return "SELECT CONVERT(INT, a.AuditID) AS AuditID, " &
                           "ISNULL(a.FullNameSnapshot, ISNULL(a.UsernameSnapshot, 'SYSTEM')) AS UserName, " &
                           "a.ActionType + CASE WHEN a.EntityType IS NULL THEN '' ELSE ' ' + a.EntityType END + " &
                           "CASE WHEN a.EntityID IS NULL THEN '' ELSE ' #' + a.EntityID END AS ActionDescription, " &
                           "CONVERT(varchar(19), a.EventAtUTC, 120) AS Timestamp " &
                           "FROM dbo.tbl_AuditTrail a " &
                           "WHERE a.AuditID > (SELECT ISNULL(MAX(AuditID), 0) FROM dbo.tbl_AuditTrail) - 1000 " &
                           "ORDER BY a.AuditID"
                Case Else
                    Return Nothing
            End Select
        End Function

    End Class
End Namespace
