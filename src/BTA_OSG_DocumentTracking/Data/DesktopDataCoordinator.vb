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
        Private ReadOnly _isDatabaseConnected As Boolean

        Public ReadOnly Property IsDatabaseConnected As Boolean
            Get
                Return _isDatabaseConnected
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
            If _isDatabaseConnected AndAlso AppStartup.DocService IsNot Nothing Then
                Dim sqlCode = RegisterDocumentInSql(docType, title, origin, dest, cab, shelf, box, gDriveUrl, initialStatus, assignedStaff, flowDir, assignedSec, deadline, lastAction, registeredByUserId, externalControlNumber)
                If sqlCode IsNot Nothing Then Return sqlCode
            End If

            ' Offline, or the SQL write failed: the cache owns the row, as before.
            EmbeddedDB.AddDocument(code, docType, title, origin, dest, cab, shelf, box, gDriveUrl, initialStatus, assignedStaff, flowDir, assignedSec, deadline, punchlist, lastAction, externalControlNumber)
            EmbeddedDB.LogAudit(registeredByName, String.Format("Registered New OSG Document [{0}] : {1} (Auto-routed to {2})", code, title, assignedSec))
            Return code
        End Function

        ''' <summary>
        ''' The connected registration. The code comes from tbl_DocumentSequences under UPDLOCK
        ''' (Services/DocumentService.vb:45), never from the local row counter, so two
        ''' workstations cannot mint the same DocCode. The document id is SQL's IDENTITY and
        ''' the cache row is a mirror of it, which is what stops one document from existing
        ''' under two keys. Returns Nothing on failure so the caller falls back to the cache.
        ''' </summary>
        Private Function RegisterDocumentInSql(docType As String, title As String, origin As String, dest As String, cab As String, shelf As String, box As String, gDriveUrl As String, initialStatus As String, assignedStaff As String, flowDir As String, assignedSec As String, deadline As String, lastAction As String, registeredByUserId As Integer, externalControlNumber As String) As String
            Try
                Dim targetDt As DateTime? = Nothing
                Dim parsedDt As DateTime
                If DateTime.TryParse(deadline, parsedDt) Then targetDt = parsedDt

                Dim doc = AppStartup.DocService.RegisterDocumentWithWorkflow(title, TypeCodeFor(docType), flowDir, origin, dest, targetDt, gDriveUrl, lastAction, registeredByUserId)
                If doc Is Nothing Then Return Nothing

                ' The desktop registers INTO a workflow: status, desk, and owner are part of the
                ' record, not defaults to be discovered later.
                Dim status = AppStartup.ReferenceDataRepo.GetStatusByCode(If(String.IsNullOrWhiteSpace(initialStatus), "RECEIVED", initialStatus))
                If status Is Nothing Then status = AppStartup.ReferenceDataRepo.GetStatusByCode("RECEIVED")
                Dim statusId As Integer = If(status IsNot Nothing, status.StatusID, 1)
                Dim effectiveSection = If(String.IsNullOrWhiteSpace(assignedSec), DocumentService.GetDefaultSectionForCategory(docType), assignedSec)
                AppStartup.DocumentRepo.UpdateWorkflowState(doc.DocumentID, statusId, effectiveSection, Nothing, lastAction, registeredByUserId)

                Dim ownerId = ResolveUserId(assignedStaff)
                If ownerId.HasValue Then
                    AppStartup.DocumentRepo.AddAssignment(New DocumentAssignment With {
                        .DocumentID = doc.DocumentID,
                        .AssignedUserID = ownerId,
                        .AssignedByUserID = registeredByUserId,
                        .AssignedAtUTC = DateTime.UtcNow,
                        .Remarks = "Registered and routed to " & effectiveSection
                    })
                End If

                AppStartup.RoutingRepo.Insert(New RoutingLog With {
                    .DocumentID = doc.DocumentID,
                    .FromStatusID = Nothing,
                    .ToStatusID = statusId,
                    .FromOffice = origin,
                    .ToOffice = dest,
                    .RoutedByUserID = registeredByUserId,
                    .RoutedAtUTC = DateTime.UtcNow,
                    .RoutingRemarks = "Document registered in OSG System."
                })

                Dim locationId = AppStartup.StorageRepo.GetOrCreateByKey(cab, shelf, box)
                AppStartup.StorageRepo.InsertMovement(New DocumentMovement With {
                    .DocumentID = doc.DocumentID,
                    .StorageLocationID = locationId,
                    .MovedByUserID = registeredByUserId,
                    .MovedAtUTC = DateTime.UtcNow,
                    .MovementReason = "Initial storage assignment."
                })
                AppStartup.DocumentRepo.UpdateStorageLocation(doc.DocumentID, locationId)

                ' Refresh the operator's grid in the same click via quick LAN read.
                EmbeddedDB.ApplySnapshots(BuildSqlSnapshot("Documents", "RoutingLogs", "Movements", "AuditTrail"))
                Return doc.DocCode
            Catch ex As Exception
                System.Diagnostics.Trace.WriteLine("SQL Server Document Register error (falling back to embedded cache): " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Scans local cache for records created while offline (PendingSync = True) and
        ''' uploads them to SQL Server, minting authoritative sequences and identities.
        ''' Must run before snapshot pulls so offline work is never lost.
        ''' </summary>
        Public Function SyncOfflineOutbox() As Integer
            If Not _isDatabaseConnected OrElse Not Program.IsDatabaseConnected Then Return 0
            If AppStartup.DocService Is Nothing Then Return 0

            Dim uploadedCount As Integer = 0
            Dim dtDocs = EmbeddedDB.DataSet.Tables("Documents")
            If dtDocs Is Nothing OrElse Not dtDocs.Columns.Contains("PendingSync") Then Return 0

            Dim pendingRows = dtDocs.Select("PendingSync = True")
            For Each row In pendingRows
                Try
                    Dim docType = row("DocType").ToString()
                    Dim title = row("Title").ToString()
                    Dim origin = row("OriginatingOffice").ToString()
                    Dim dest = row("DestinationOffice").ToString()
                    Dim cab = row("CabinetID").ToString()
                    Dim shelf = row("ShelfNo").ToString()
                    Dim box = row("BoxCode").ToString()
                    Dim gDrive = row("GDriveURL").ToString()
                    Dim initialStatus = row("CurrentStatus").ToString()
                    Dim assigned = row("AssignedStaff").ToString()
                    Dim flowDir = row("FlowDirection").ToString()
                    Dim assignedSec = row("AssignedSection").ToString()
                    Dim deadline = row("TargetDeadlineUTC").ToString()
                    Dim lastAction = row("LastActionTaken").ToString()
                    Dim extControl = If(row.Table.Columns.Contains("ExternalControlNumber"), row("ExternalControlNumber").ToString(), "")

                    Dim authoritativeCode = RegisterDocumentInSql(docType, title, origin, dest, cab, shelf, box, gDrive, initialStatus, assigned, flowDir, assignedSec, deadline, lastAction, 1, extControl)
                    If authoritativeCode IsNot Nothing Then
                        row("PendingSync") = False
                        uploadedCount += 1
                        Try
                            dtDocs.Rows.Remove(row)
                        Catch
                        End Try
                    End If
                Catch ex As Exception
                    System.Diagnostics.Trace.WriteLine("Failed to replay offline document to SQL: " & ex.Message)
                End Try
            Next

            Return uploadedCount
        End Function

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
            If _isDatabaseConnected AndAlso AppStartup.RoutingService IsNot Nothing Then
                If TryRunSql("RouteDocument", Sub()
                                                   Dim status = AppStartup.ReferenceDataRepo.GetStatusByCode(If(String.IsNullOrWhiteSpace(action), "ROUTED", action))
                                                   Dim statusId As Integer = If(status IsNot Nothing, status.StatusID, 1)

                                                   AppStartup.DocumentRepo.UpdateWorkflowState(docId, statusId, toOffice, Nothing, "Routed to " & toOffice & ": " & action, routedByUserId, destinationOffice:=toOffice, originOffice:=fromOffice)

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
                                                   AppStartup.RoutingService.AddRoutingLog(log)
                                               End Sub) Then
                    PullAfterWrite("Documents", "RoutingLogs", "AuditTrail")
                    Return
                End If
            End If
            EmbeddedDB.AddRoutingLog(docId, fromOffice, toOffice, routedByName, action, remarks)
            EmbeddedDB.LogAudit(routedByName, String.Format("Routed Doc #{0} from {1} to {2}: {3}", docId, fromOffice, toOffice, action))
        End Sub

        Public Sub ApplyDirective(docId As Integer, directiveText As String, assignedTo As String, notes As String, staffName As String, staffUserId As Integer)
            If _isDatabaseConnected AndAlso AppStartup.DirectiveService IsNot Nothing Then
                ' The directive type stays hardcoded for now; mapping the UI choice onto
                ' tbl_DirectiveTypes.DirectiveCode is a separate concern.
                If TryRunSql("ApplyDirective", Sub() AppStartup.DirectiveService.IssueDirective(docId, 1, directiveText, notes, staffUserId)) Then
                    PullAfterWrite("Directives", "Documents", "AuditTrail")
                    Return
                End If
            End If
            EmbeddedDB.AddDirective(docId, directiveText, assignedTo, notes, staffName)
            EmbeddedDB.LogAudit(staffName, String.Format("Applied SG Directive [{0}] to Doc ID #{1}", directiveText, docId))
        End Sub

        Public Sub MoveStorage(docId As Integer, fromLoc As String, toLoc As String, reason As String, movedByName As String, movedByUserId As Integer)
            If _isDatabaseConnected AndAlso AppStartup.StorageService IsNot Nothing Then
                ' The whole point of a move is which box the folder went to, so resolve the real
                ' landmark first; MoveDocument used to be handed a hardcoded location id.
                Dim parts = If(toLoc, "").Split("|"c)
                Dim cab = If(parts.Length > 0, parts(0), "")
                Dim shelf = If(parts.Length > 1, parts(1), "")
                Dim box = If(parts.Length > 2, parts(2), "")
                Dim locationId As Integer = 0
                If TryRunSql("Storage lookup", Sub() locationId = AppStartup.StorageRepo.GetOrCreateByKey(cab, shelf, box)) Then
                    If TryRunSql("MoveStorage", Sub() AppStartup.StorageService.MoveDocument(docId, locationId, movedByUserId, reason)) Then
                        PullAfterWrite("Documents", "Movements", "AuditTrail")
                        Return
                    End If
                End If
            End If
            EmbeddedDB.AddMovementLog(docId, fromLoc, toLoc, movedByName, reason)
            EmbeddedDB.LogAudit(movedByName, String.Format("Transferred Doc #{0} physical location from {1} to {2}", docId, fromLoc, toLoc))
        End Sub

        Public Sub RequestRevision(docId As Integer, punchlistNotes As String, returnSection As String, staffName As String, staffUserId As Integer)
            If _isDatabaseConnected AndAlso AppStartup.DirectiveService IsNot Nothing Then
                If TryRunSql("RequestRevision", Sub() AppStartup.DirectiveService.RequestRevision(docId, punchlistNotes, staffUserId, returnSection)) Then
                    PullAfterWrite("Documents", "Directives", "RoutingLogs", "AuditTrail")
                    Return
                End If
            End If
            EmbeddedDB.RequestRevision(docId, punchlistNotes, returnSection, staffName)
        End Sub

        Public Sub ResubmitDocument(docId As Integer, staffName As String, notes As String, staffUserId As Integer)
            If _isDatabaseConnected AndAlso AppStartup.RoutingService IsNot Nothing Then
                If TryRunSql("ResubmitDocument", Sub() AppStartup.RoutingService.ResubmitDocument(docId, staffUserId, notes)) Then
                    PullAfterWrite("Documents", "RoutingLogs", "Movements", "AuditTrail")
                    Return
                End If
            End If
            EmbeddedDB.ResubmitDocument(docId, staffName, notes)
        End Sub

        Public Sub ApproveDocument(docId As Integer, staffName As String, notes As String, staffUserId As Integer)
            If _isDatabaseConnected AndAlso AppStartup.RoutingService IsNot Nothing Then
                If TryRunSql("ApproveDocument", Sub() AppStartup.RoutingService.ApproveDocument(docId, staffUserId, notes)) Then
                    PullAfterWrite("Documents", "RoutingLogs", "AuditTrail")
                    Return
                End If
            End If
            EmbeddedDB.ApproveDocument(docId, staffName, notes)
        End Sub

        Public Sub ReleaseDocument(docId As Integer, staffName As String, notes As String, staffUserId As Integer)
            If _isDatabaseConnected AndAlso AppStartup.RoutingService IsNot Nothing Then
                If TryRunSql("ReleaseDocument", Sub() AppStartup.RoutingService.ReleaseDocument(docId, staffUserId, notes)) Then
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
            If _isDatabaseConnected AndAlso AppStartup.UserRepo IsNot Nothing Then
                Try
                    Dim existing = AppStartup.UserRepo.GetByCardPublicID(cleanUid)
                    If existing IsNot Nothing Then
                        ' Re-enrolment: update the person and their flags instead of creating a
                        ' second account (tbl_RfidCards.CardPublicID is UNIQUE).
                        existing.FullName = fullName
                        existing.Office = office
                        existing.IsActive = True
                        existing.CanRoute = canRoute
                        existing.CanMove = canMove
                        existing.CanSoftCopy = canSoftCopy
                        AppStartup.UserRepo.Update(existing, enrolledByUserId)
                        AppStartup.UserRepo.AssignRole(existing.UserID, RoleCodeFor(role), enrolledByUserId)
                        PullAfterWrite("Users", "AuditTrail")
                        Return True
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
                        userId = AppStartup.UserRepo.Insert(newUser, enrolledByUserId)
                    Catch dup As Microsoft.Data.SqlClient.SqlException
                        ' Username is UNIQUE too; two staff with the same name get the badge suffix.
                        newUser.Username = SanitizeUsername(fullName) & "_" & Right(cleanUid, 4).ToLowerInvariant()
                        userId = AppStartup.UserRepo.Insert(newUser, enrolledByUserId)
                    End Try

                    AppStartup.CardService.IssueCard(userId, cleanUid, "Enrolled from desktop Admin", enrolledByUserId)
                    AppStartup.UserRepo.AssignRole(userId, RoleCodeFor(role), enrolledByUserId)

                    PullAfterWrite("Users", "AuditTrail")
                    Return True
                Catch ex As Exception
                    System.Diagnostics.Trace.WriteLine("SQL Server user enrolment error (falling back to embedded cache): " & ex.Message)
                End Try
            End If

            EmbeddedDB.AddUser(cleanUid, fullName, role, office, canRoute, canMove, canSoftCopy)
            EmbeddedDB.LogAudit(enrolledByName, String.Format("Registered/Updated User [{0}] Role: {1} (Desk: {2}) Privileges: [Route:{3}, Move:{4}, SoftCopy:{5}] RFID: ****{6}", fullName, role, office, canRoute, canMove, canSoftCopy, Right(cleanUid, 4)))
            Return False
        End Function

        ''' <summary>
        ''' Desktop role or section name to SQL role code. The SQL role names are what the
        ''' Users mirror projects back into the cache Role column, so this map is what keeps
        ''' the section desk logic (EmbeddedDB.GetVisibleDocuments) working after a sync.
        ''' </summary>
        Private Shared Function RoleCodeFor(role As String) As String
            Select Case If(role, "").Trim().ToUpperInvariant()
                Case "SECRETARY-GENERAL", "SG" : Return "SG"
                Case "SYSTEM ADMINISTRATOR", "SYSADMIN" : Return "SYSADMIN"
                Case "OSG CHIEF" : Return "OSG_CHIEF"
                Case "RECORDS SECTION" : Return "RECORDS"
                Case "SECRETARIAT" : Return "SECRETARIAT"
                Case "LEGISLATIVE SECTION" : Return "LEGISLATIVE"
                Case "FINANCE SECTION" : Return "FINANCE"
                Case "TRAVEL SECTION" : Return "TRAVEL"
                Case Else : Return "ADMIN_STAFF"
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
        ''' </summary>
        Private Function TryRunSql(label As String, action As Action) As Boolean
            Try
                action()
                Return True
            Catch ex As Exception
                System.Diagnostics.Trace.WriteLine("SQL Server " & label & " error (falling back to embedded cache): " & ex.Message)
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
            If Not _isDatabaseConnected OrElse Not Program.IsDatabaseConnected Then Return Nothing
            If AppStartup.Settings Is Nothing Then Return Nothing

            Dim names = If(tables Is Nothing OrElse tables.Length = 0,
                           New String() {"Documents", "Users", "Directives", "RoutingLogs", "Movements", "AuditTrail"},
                           tables)
            Dim result As New Dictionary(Of String, DataTable)()
            Try
                ' Connect with a 5s budget so an unreachable host skips a tick promptly.
                Dim builder As New Microsoft.Data.SqlClient.SqlConnectionStringBuilder(AppStartup.Settings.DatabaseSettings.ConnectionString) With {
                    .ConnectTimeout = 5
                }
                Using conn As New Microsoft.Data.SqlClient.SqlConnection(builder.ConnectionString)
                    conn.Open()
                    For Each name In names
                        Dim sql = SqlFor(name)
                        If sql Is Nothing Then Continue For
                        Try
                            Using cmd As New Microsoft.Data.SqlClient.SqlCommand(sql, conn)
                                Using reader = cmd.ExecuteReader()
                                    Dim table = EmbeddedDB.DataSet.Tables(name).Clone()
                                    table.Load(reader)
                                    result(name) = table
                                End Using
                            End Using
                        Catch ex As Exception
                            System.Diagnostics.Trace.WriteLine("SQL snapshot read failed for " & name & " (table left as cached): " & ex.Message)
                        End Try
                    Next
                End Using
            Catch ex As Exception
                System.Diagnostics.Trace.WriteLine("SQL snapshot connection failed: " & ex.Message)
                Return Nothing
            End Try
            Return result
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
                           "ISNULL(d.ExternalControlNumber, '') AS ExternalControlNumber " &
                           "FROM dbo.tbl_Documents d " &
                           "LEFT JOIN dbo.tbl_DocumentTypes dt ON dt.DocumentTypeID = d.DocumentTypeID " &
                           "LEFT JOIN dbo.tbl_DocumentStatuses st ON st.StatusID = d.StatusID " &
                           "LEFT JOIN dbo.tbl_StorageLocations sl ON sl.StorageLocationID = d.CurrentStorageLocationID " &
                           "OUTER APPLY (SELECT TOP 1 u.FullName FROM dbo.tbl_DocumentAssignments da " &
                           "  JOIN dbo.tbl_Users u ON u.UserID = da.AssignedUserID " &
                           "  WHERE da.DocumentID = d.DocumentID ORDER BY da.AssignmentID DESC) a " &
                           "WHERE d.IsDeleted = 0 ORDER BY d.DocumentID"
                Case "Users"
                    Return "SELECT u.UserID, ISNULL(c.CardPublicID, '') AS RFID_UID, ISNULL(u.FullName, '') AS FullName, " &
                           "ISNULL(ro.RoleName, ISNULL(u.Office, 'Administrative Staff')) AS Role, ISNULL(u.Office, '') AS Office, " &
                           "ISNULL(u.CanRoute, 1) AS CanRoute, ISNULL(u.CanMove, 1) AS CanMove, ISNULL(u.CanSoftCopy, 1) AS CanSoftCopy, ISNULL(u.IsActive, 1) AS IsActive " &
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
