Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Data

Namespace BTA_OSG
    ' SQL Server mirror: reads the shared database into cache-shaped snapshot tables
    ' (BuildSqlSnapshot and its projections), normalizes UTC log times to wall clock at
    ' the mirror boundary, and reports the workstation heartbeat. Split from
    ' DesktopDataCoordinator.vb (same partial class, zero behavior change) so the
    ' dispatch file stays under the 1k-line review threshold.
    Partial Public Class DesktopDataCoordinator

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
                    _sqlAvailable = True
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
                                        NormalizeMirroredTimes(name, table)
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
                _sqlAvailable = False
                System.Diagnostics.Trace.WriteLine("SQL snapshot connection failed: " & ex.Message)
                Return Nothing
            End Try
            Return result
        End Function

        ' SQL stores these four logs in UTC; the cache, grids, detail timeline, and printed
        ' slip all show wall-clock time. Convert once at the mirror boundary so offline and
        ' mirrored rows share one basis. Heartbeat stays UTC: its staleness check compares it
        ' with DateTime.UtcNow.
        Friend Shared Sub NormalizeMirroredTimes(tableName As String, table As DataTable)
            If table Is Nothing Then Return
            Select Case tableName
                Case "Directives", "RoutingLogs", "Movements", "AuditTrail"
                Case Else
                    Return
            End Select
            Const columnName As String = "Timestamp"
            If Not table.Columns.Contains(columnName) Then Return
            For Each row As DataRow In table.Rows
                If IsDBNull(row(columnName)) Then Continue For
                Dim parsed As DateTime
                If DateTime.TryParse(row(columnName).ToString(), parsed) Then
                    row(columnName) = DateTime.SpecifyKind(parsed, DateTimeKind.Utc).ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss")
                End If
            Next
        End Sub

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
                           "ISNULL(d.ExternalControlNumber, '') AS ExternalControlNumber, ISNULL(d.RequesterGender, '') AS RequesterGender, d.RowVersion " &
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
                    ' A voided document's operational history stays in SQL and the audit trail;
                    ' the active grids must not keep showing directives for a row the registry
                    ' no longer lists.
                    Return "SELECT ad.DirectiveID, ad.DocumentID, ISNULL(ad.DirectiveText, '') AS SGDirective, " &
                           "ISNULL(u.FullName, '') AS AssignedTo, ISNULL(ad.Remarks, '') AS Notes, " &
                           "ISNULL(u.FullName, 'SYSTEM') AS LogUser, CONVERT(varchar(19), ad.IssuedAtUTC, 120) AS Timestamp " &
                           "FROM dbo.tbl_ActionDirectives ad " &
                           "JOIN dbo.tbl_Documents doc ON doc.DocumentID = ad.DocumentID AND doc.IsDeleted = 0 " &
                           "LEFT JOIN dbo.tbl_Users u ON u.UserID = ad.IssuedByUserID " &
                           "ORDER BY ad.DirectiveID"
                Case "RoutingLogs"
                    Return "SELECT r.RoutingLogID AS RoutingID, r.DocumentID, ISNULL(r.FromOffice, '') AS FromOffice, " &
                           "ISNULL(r.ToOffice, '') AS ToOffice, ISNULL(u.FullName, 'SYSTEM') AS RoutedBy, " &
                           "ISNULL(st.StatusCode, '') AS ActionTaken, ISNULL(r.RoutingRemarks, '') AS Remarks, " &
                           "CONVERT(varchar(19), r.RoutedAtUTC, 120) AS Timestamp " &
                           "FROM dbo.tbl_RoutingLogs r " &
                           "JOIN dbo.tbl_Documents doc ON doc.DocumentID = r.DocumentID AND doc.IsDeleted = 0 " &
                           "LEFT JOIN dbo.tbl_Users u ON u.UserID = r.RoutedByUserID " &
                           "LEFT JOIN dbo.tbl_DocumentStatuses st ON st.StatusID = r.ToStatusID " &
                           "ORDER BY r.RoutingLogID"
                Case "Movements"
                    Return "SELECT m.MovementID, m.DocumentID, '' AS FromLocation, " &
                           "ISNULL(sl.CabinetID, '') + '|' + ISNULL(sl.ShelfNo, '') + '|' + ISNULL(sl.BoxCode, '') AS ToLocation, " &
                           "ISNULL(u.FullName, 'SYSTEM') AS MovedBy, ISNULL(m.MovementReason, '') AS Reason, " &
                           "CONVERT(varchar(19), m.MovedAtUTC, 120) AS Timestamp " &
                           "FROM dbo.tbl_DocumentMovements m " &
                           "JOIN dbo.tbl_Documents doc ON doc.DocumentID = m.DocumentID AND doc.IsDeleted = 0 " &
                           "LEFT JOIN dbo.tbl_Users u ON u.UserID = m.MovedByUserID " &
                           "LEFT JOIN dbo.tbl_StorageLocations sl ON sl.StorageLocationID = m.StorageLocationID " &
                           "ORDER BY m.MovementID"
                Case "AuditTrail"
                    ' The audit trail is append-only; mirror the newest 1000 events in chronological order.
                    Return "SELECT CONVERT(INT, a.AuditID) AS AuditID, " &
                           "ISNULL(a.FullNameSnapshot, ISNULL(a.UsernameSnapshot, 'SYSTEM')) AS UserName, " &
                           "a.ActionType + CASE WHEN a.EntityType IS NULL THEN '' ELSE ' ' + a.EntityType END + " &
                           "CASE WHEN a.EntityID IS NULL THEN '' ELSE ' #' + a.EntityID END + " &
                           "CASE WHEN a.ActionType = 'DOCUMENT_VOIDED' AND a.DocumentCode IS NOT NULL THEN ' [' + a.DocumentCode + ']' ELSE '' END + " &
                           "CASE WHEN a.ActionType = 'DOCUMENT_VOIDED' AND ISJSON(a.NewValuesJson) = 1 " &
                           "THEN ' | ' + ISNULL(JSON_VALUE(a.NewValuesJson, '$.VoidReason'), '') ELSE '' END AS ActionDescription, " &
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
