Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Data

Namespace BTA_OSG
    ' Offline outbox replay: carries records the workstation created or edited while
    ' offline into SQL Server once connectivity returns (SyncOfflineOutbox and the
    ' Replay* family), resolving local autoincrement user ids to SQL ids and remapping
    ' child rows to the replayed document. Split from DesktopDataCoordinator.vb (same
    ' partial class, zero behavior change) so the dispatch file stays under the
    ' 1k-line review threshold.
    Partial Public Class DesktopDataCoordinator

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
                            .DateReceived = If(row.Table.Columns.Contains("DateReceived"), row("DateReceived").ToString(), ""),
                            .LastAction = row("LastActionTaken").ToString(),
                            .ExternalControlNumber = If(row.Table.Columns.Contains("ExternalControlNumber"), row("ExternalControlNumber").ToString(), ""),
                            .RequesterGender = If(row.Table.Columns.Contains("RequesterGender"), row("RequesterGender").ToString(), ""),
                            .PreferredDocCode = row("DocCode").ToString()
                        }
                    End SyncLock

                    ' A duplicate offline registration (the same portal submission imported on
                    ' two seats, or a retry after a half-written row) must resolve to the
                    ' existing server document instead of dying forever on the filtered
                    ' unique index and stranding its child rows in the outbox.
                    If fields.ExternalControlNumber.Length > 0 AndAlso AppStartup.DocumentRepo IsNot Nothing Then
                        Dim existingByEcn = AppStartup.DocumentRepo.GetByExternalControlNumber(fields.ExternalControlNumber)
                        If existingByEcn IsNot Nothing Then
                            SyncLock EmbeddedDB.SyncRoot
                                RemapOfflineChildRecords(oldLocalId, existingByEcn.DocumentID)
                                row("PendingSync") = False
                                If isNewCol Then row("IsNewOfflineRecord") = False
                                Try
                                    dtDocs.Rows.Remove(row)
                                Catch ex2 As Exception
                                    System.Diagnostics.Trace.WriteLine("Replay: could not remove merged offline row " & oldLocalId.ToString() & ": " & ex2.Message)
                                End Try
                            End SyncLock
                            EmbeddedDB.Save()
                            uploadedCount += 1
                            Continue For
                        End If
                    End If

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
                            Catch removeEx As Exception
                                ' The SQL row is committed; a failed local removal only leaves a
                                ' ghost duplicate until the next snapshot pull reconciles it.
                                System.Diagnostics.Trace.WriteLine("Replay: could not remove replayed row " & oldLocalId.ToString() & ": " & removeEx.Message)
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
                    Dim assignedStaff As String = ""
                    Dim modifiedBy As Integer?
                    Dim rowVersion As Byte() = Nothing
                    SyncLock EmbeddedDB.SyncRoot
                        assignedSec = row("AssignedSection").ToString()
                        punchlist = row("RevisionPunchlist").ToString()
                        lastAction = row("LastActionTaken").ToString()
                        destOffice = row("DestinationOffice").ToString()
                        originOffice = row("OriginatingOffice").ToString()
                        If dtDocs.Columns.Contains("AssignedStaff") AndAlso Not IsDBNull(row("AssignedStaff")) Then
                            assignedStaff = row("AssignedStaff").ToString()
                        End If
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
                        ' The cache row's AssignedStaff is the desk the offline directive chose.
                        ' Without this, the assignment exists only in the XML cache and silently
                        ' reverts on the next snapshot pull (the mirror derives the column from
                        ' tbl_DocumentAssignments). The mirror keeps the latest row per document,
                        ' so re-recording an unchanged assignment is harmless.
                        If assignedStaff.Trim().Length > 0 AndAlso AppStartup.DocumentRepo IsNot Nothing Then
                            Dim assigneeId = ResolveUserId(assignedStaff)
                            If assigneeId.HasValue Then
                                Try
                                    AppStartup.DocumentRepo.AddAssignment(New DocumentAssignment With {
                                        .DocumentID = docId,
                                        .AssignedUserID = assigneeId.Value,
                                        .AssignedByUserID = If(modifiedBy.HasValue, modifiedBy.Value, 0),
                                        .AssignedAtUTC = DateTime.UtcNow,
                                        .Remarks = "Replayed offline assignment"
                                    })
                                Catch assignEx As Exception
                                    ' The state replay succeeded; a failed assignment row must not
                                    ' flip the document back to pending and replay the status twice.
                                    System.Diagnostics.Trace.WriteLine("Replay assignment insert failed for doc " & docId.ToString() & ": " & assignEx.Message)
                                End Try
                            End If
                        End If
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
        ''' A child row may only replay once its parent document is server truth. When the
        ''' parent is still an offline-created row (PendingSync and IsNewOfflineRecord), its
        ''' DocumentID is a LOCAL autoincrement id: GetById would match whichever server
        ''' document happens to hold that number and attach the directive, routing log, or
        ''' movement to a stranger. Children of mirrored documents (or rows already remapped
        ''' by RemapOfflineChildRecords) carry real SQL ids and pass through.
        ''' </summary>
        Private Function ParentDocumentReady(docId As Integer) As Boolean
            SyncLock EmbeddedDB.SyncRoot
                Dim dtDocs = EmbeddedDB.DataSet.Tables("Documents")
                If dtDocs IsNot Nothing Then
                    Dim rows = dtDocs.Select(String.Format("DocumentID = {0}", docId))
                    If rows.Length > 0 AndAlso dtDocs.Columns.Contains("PendingSync") AndAlso CBool(rows(0)("PendingSync")) _
                       AndAlso dtDocs.Columns.Contains("IsNewOfflineRecord") AndAlso Not IsDBNull(rows(0)("IsNewOfflineRecord")) AndAlso CBool(rows(0)("IsNewOfflineRecord")) Then
                        Return False
                    End If
                End If
            End SyncLock
            Return AppStartup.DocumentRepo.GetById(docId) IsNot Nothing
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
                    Dim issuedAtUtc As DateTime
                    SyncLock EmbeddedDB.SyncRoot
                        docId = CInt(row("DocumentID"))
                        directiveText = row("SGDirective").ToString()
                        notes = row("Notes").ToString()
                        issuedBy = ReplayUserId(row, "IssuedByUserID")
                        code = If(row.Table.Columns.Contains("DirectiveCode") AndAlso Not IsDBNull(row("DirectiveCode")), row("DirectiveCode").ToString(), "")
                        issuedAtUtc = CachedEventUtc(row, "Timestamp")
                    End SyncLock
                    ' The parent document must already exist in SQL: inserting against the
                    ' local id would attach the directive to whatever document holds that id
                    ' on the server. Rows stay queued until their document replays.
                    If Not ParentDocumentReady(docId) Then Continue For
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
                        .IssuedAtUTC = issuedAtUtc,
                        .IsActive = True,
                        .Remarks = notes
                    }
                    ' Insert and audit are one fact, same as the connected IssueDirective: an
                    ' audit failure after a bare commit would leave the row pending and the
                    ' next tick would file the directive a second time.
                    Using conn = AppStartup.ConnectionFactory.CreateConnection()
                        Using tx = conn.BeginTransaction()
                            Try
                                AppStartup.DirectiveRepo.Insert(directive, tx)
                                If AppStartup.AuditService IsNot Nothing Then
                                    AppStartup.AuditService.LogEvent("DIRECTIVE_ADDED", "Directive", docId.ToString(), Nothing, Nothing, Nothing, True, Nothing, tx)
                                End If
                                tx.Commit()
                            Catch
                                tx.Rollback()
                                Throw
                            End Try
                        End Using
                    End Using
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
                        Dim routedAtUtc As DateTime
                        SyncLock EmbeddedDB.SyncRoot
                            docId = CInt(row("DocumentID"))
                            fromOffice = row("FromOffice").ToString()
                            toOffice = row("ToOffice").ToString()
                            remarks = row("Remarks").ToString()
                            routedBy = ReplayUserId(row, "RoutedByUserID")
                            routedAtUtc = CachedEventUtc(row, "Timestamp")
                        End SyncLock
                        If Not ParentDocumentReady(docId) Then Continue For
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
                            .RoutedAtUTC = routedAtUtc
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
                        Dim movedAtUtc As DateTime
                        SyncLock EmbeddedDB.SyncRoot
                            docId = CInt(row("DocumentID"))
                            toLoc = row("ToLocation").ToString()
                            movedBy = ReplayUserId(row, "MovedByUserID")
                            movedAtUtc = CachedEventUtc(row, "Timestamp")
                        End SyncLock
                        If Not ParentDocumentReady(docId) Then Continue For
                        If Not movedBy.HasValue Then
                            WarnReplayOnce("movements", "an offline storage movement could not be attributed to a known user")
                            Continue For
                        End If
                        Dim parts = toLoc.Split(New Char() {"/"c, "|"c})
                        Dim cab = If(parts.Length > 0, parts(0).Trim(), "")
                        Dim shelf = If(parts.Length > 1, parts(1).Trim(), "")
                        Dim box = If(parts.Length > 2, parts(2).Trim(), "")

                        Dim locationId = AppStartup.StorageRepo.GetOrCreateByKey(cab, shelf, box)

                        AppStartup.StorageService.MoveDocument(docId, locationId, movedBy.Value, reason, movedAtUtc:=movedAtUtc)
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

        ' Offline log rows stamp the wall clock (EmbeddedDB uses DateTime.Now); SQL stores UTC.
        ' Replay carries the moment the action happened instead of the moment the seat synced.
        Friend Shared Function CachedEventUtc(row As DataRow, columnName As String) As DateTime
            Dim parsed As DateTime
            If row IsNot Nothing AndAlso row.Table.Columns.Contains(columnName) AndAlso
               Not IsDBNull(row(columnName)) AndAlso DateTime.TryParse(row(columnName).ToString(), parsed) Then
                Return DateTime.SpecifyKind(parsed, DateTimeKind.Local).ToUniversalTime()
            End If
            Return DateTime.UtcNow
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
                    Dim eventAt As DateTime
                    SyncLock EmbeddedDB.SyncRoot
                        actionDesc = row("ActionDescription").ToString()
                        userId = ReplayUserId(row, "UserID")
                        ' Keep the original action type and actor so the SQL audit trail reads
                        ' as the events that happened, not as one opaque OFFLINE_SYNC blob.
                        originalType = If(row.Table.Columns.Contains("ActionType") AndAlso Not IsDBNull(row("ActionType")), row("ActionType").ToString(), "")
                        actor = If(row.Table.Columns.Contains("UserName") AndAlso Not IsDBNull(row("UserName")), row("UserName").ToString(), "SYSTEM")
                        auditId = row("AuditID").ToString()
                        ' The cache Timestamp is the wall-clock moment the action happened
                        ' (LogAudit writes local time); replaying DateTime.UtcNow would show
                        ' the sync batch hour in the forensic timeline instead.
                        eventAt = CachedEventUtc(row, "Timestamp")
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
                        .EventAtUTC = eventAt
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
    End Class
End Namespace
