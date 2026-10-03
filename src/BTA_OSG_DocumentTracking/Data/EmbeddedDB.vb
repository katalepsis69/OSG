Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.IO

Namespace BTA_OSG
    ''' <summary>
    ''' Embedded Relational Database Engine (DataSet + XML Persistence).
    ''' Provides zero-configuration local storage and offline/demo fallback per Path A specification.
    ''' </summary>
    Public Class EmbeddedDB
        ' Friend so the test assembly can point the store at a scratch file; production
        ' always uses the exe folder.
        Friend Shared DbPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "bta_osg_db.xml")
        Public Shared DataSet As New DataSet("BTA_OSG_DB")

        ' Set when WriteXml fails (disk full, file locked): callers that only fire-and-forget
        ' mutations can surface it instead of the offline work silently stopping persisting.
        Public Shared LastPersistenceError As String = ""
        Public Shared Event PersistenceFailed(message As String)

        Public Shared Sub Initialize()
            DataSet.Clear()
            DataSet.Tables.Clear()

            CreateTables()

            If File.Exists(DbPath) Then
                Try
                    ' IgnoreSchema loads the persisted rows into the freshly created (current)
                    ' schema: columns the file predates take their defaults, columns the file
                    ' has that the code dropped are ignored. This is what makes offline work
                    ' survive a restart instead of the cache being write-only.
                    DataSet.ReadXml(DbPath, XmlReadMode.IgnoreSchema)
                    ResyncAutoIncrementCounters()
                    Return
                Catch ex As Exception
                    System.Diagnostics.Trace.TraceError("EmbeddedDB cache load failed (starting empty): " & ex.Message)
                End Try
            End If
            Save()
        End Sub

        ' ReadXml loads explicit PK values without advancing the AutoIncrement counters, so a
        ' fresh row could collide with a loaded one. Point each identity column past the max.
        Private Shared Sub ResyncAutoIncrementCounters()
            For Each table As DataTable In DataSet.Tables
                For Each pk As DataColumn In table.PrimaryKey
                    If Not pk.AutoIncrement Then Continue For
                    Dim maxValue As Long = 0
                    For Each row As DataRow In table.Rows
                        If Not IsDBNull(row(pk)) Then
                            Dim v = Convert.ToInt64(row(pk))
                            If v > maxValue Then maxValue = v
                        End If
                    Next
                    pk.AutoIncrementSeed = maxValue + pk.AutoIncrementStep
                Next
            Next
        End Sub

        Public Shared Sub EnsureInitialized()
            If DataSet.Tables.Count = 0 Then
                Initialize()
            End If
        End Sub

        Public Shared ReadOnly Property SyncRoot As Object
            Get
                Return _syncLock
            End Get
        End Property

        Private Shared ReadOnly _syncLock As New Object()
        Private Shared ReadOnly _flushLock As New Object()
        Private Shared _flushTimer As System.Threading.Timer
        Private Const FlushDelayMs As Integer = 1500

        ''' <summary>
        ''' Immediate, synchronous persistence. Reserved for explicit durability points
        ''' (first-run seeding, app exit, self-check); routine mutations use MarkDirty so a
        ''' multi-step action pays one disk write instead of one write per table touched.
        ''' </summary>
        Public Shared Sub Save()
            CancelPendingFlush()
            SyncLock _syncLock
                Try
                    DataSet.WriteXml(DbPath, XmlWriteMode.WriteSchema)
                    LastPersistenceError = ""
                Catch ex As Exception
                    ' A silently lost flush would strand PendingSync rows in the outbox with
                    ' no evidence anywhere, so the failure must reach the trace log AND the
                    ' operator surface (FormMain listens for PersistenceFailed).
                    LastPersistenceError = ex.Message
                    System.Diagnostics.Trace.TraceError("EmbeddedDB.Save failed: " & ex.Message)
                    RaiseEvent PersistenceFailed(ex.Message)
                End Try
            End SyncLock
        End Sub

        ''' <summary>
        ''' Merges one SQL Server snapshot into its cache table. SQL Server is the source of
        ''' truth whenever it is reachable. Rows are updated in place rather than cleared and reloaded.
        ''' Rows marked with PendingSync = True are strictly preserved so offline work is never wiped.
        ''' </summary>
        Public Shared Sub ApplySnapshot(tableName As String, snapshot As DataTable, Optional persist As Boolean = True)
            If snapshot Is Nothing OrElse snapshot.Rows.Count = 0 Then Return
            SyncLock _syncLock
                EnsureInitialized()
                If Not DataSet.Tables.Contains(tableName) Then Return
                Dim target = DataSet.Tables(tableName)
                If target.PrimaryKey Is Nothing OrElse target.PrimaryKey.Length <> 1 Then Return
                Dim pk = target.PrimaryKey(0)
                If Not snapshot.Columns.Contains(pk.ColumnName) Then Return

                Dim incoming As New HashSet(Of Object)()
                For Each source As DataRow In snapshot.Rows
                    incoming.Add(source(pk.ColumnName))
                Next

                ' Merge row by row instead of Load(OverwriteChanges): a PK match on a
                ' PendingSync row must keep the offline data, not be overwritten by the
                ' mirrored SQL row that happens to carry the same local id.
                For Each source As DataRow In snapshot.Rows
                    Dim existing = target.Rows.Find(source(pk.ColumnName))
                    If existing IsNot Nothing AndAlso IsPendingRow(target, existing) Then Continue For
                    If existing Is Nothing Then
                        Dim added = target.NewRow()
                        CopySnapshotValues(source, added, target)
                        target.Rows.Add(added)
                    Else
                        CopySnapshotValues(source, existing, target)
                    End If
                Next

                If incoming.Count > 0 Then
                    Dim doomed As New List(Of DataRow)()
                    For Each row As DataRow In target.Rows
                        If Not IsPendingRow(target, row) AndAlso Not incoming.Contains(row(pk)) Then
                            ' Protect initial seeded local documents (IDs 1-4) from deletion if incoming is a partial set
                            If tableName.Equals("Documents", StringComparison.OrdinalIgnoreCase) Then
                                Dim idVal = If(pk.ColumnName = "DocumentID", Convert.ToInt32(row(pk)), 0)
                                If idVal <= 4 AndAlso Not incoming.Contains(row(pk)) Then Continue For
                            End If
                            doomed.Add(row)
                        End If
                    Next
                    For Each row As DataRow In doomed
                        target.Rows.Remove(row)
                    Next
                End If

                If persist Then MarkDirty()
            End SyncLock
        End Sub

        Private Shared Function IsPendingRow(table As DataTable, row As DataRow) As Boolean
            If Not table.Columns.Contains("PendingSync") Then Return False
            If IsDBNull(row("PendingSync")) Then Return False
            Return CBool(row("PendingSync"))
        End Function

        Private Shared Sub CopySnapshotValues(source As DataRow, targetRow As DataRow, target As DataTable)
            For Each col As DataColumn In target.Columns
                If Not source.Table.Columns.Contains(col.ColumnName) Then Continue For
                targetRow(col) = source(col.ColumnName)
            Next
        End Sub

        Public Shared Sub ApplySnapshots(snapshots As Dictionary(Of String, DataTable), Optional persist As Boolean = True)
            If snapshots Is Nothing OrElse snapshots.Count = 0 Then Return
            For Each pair In snapshots
                ApplySnapshot(pair.Key, pair.Value, persist)
            Next
        End Sub

        ''' <summary>
        ''' Schedules a batched write after a short idle window. Safe to call on every mutation;
        ''' the timer keeps sliding while changes keep coming, and app exit flushes via Save().
        ''' Never holds _flushLock and _syncLock at the same time, so no lock-order inversion.
        ''' </summary>
        Private Shared Sub MarkDirty()
            SyncLock _flushLock
                If _flushTimer Is Nothing Then
                    _flushTimer = New System.Threading.Timer(AddressOf FlushTimerTick, Nothing, FlushDelayMs, System.Threading.Timeout.Infinite)
                Else
                    _flushTimer.Change(FlushDelayMs, System.Threading.Timeout.Infinite)
                End If
            End SyncLock
        End Sub

        Private Shared Sub CancelPendingFlush()
            SyncLock _flushLock
                If _flushTimer IsNot Nothing Then
                    _flushTimer.Change(System.Threading.Timeout.Infinite, System.Threading.Timeout.Infinite)
                End If
            End SyncLock
        End Sub

        Private Shared Sub FlushTimerTick(state As Object)
            Save()
        End Sub

        Private Shared Sub CreateTables()
            ' tbl_Users
            Dim dtUsers As New DataTable("Users")
            dtUsers.Columns.Add("UserID", GetType(Integer)).AutoIncrement = True
            dtUsers.Columns("UserID").AutoIncrementSeed = 1
            dtUsers.Columns("UserID").AutoIncrementStep = 1
            dtUsers.Columns.Add("RFID_UID", GetType(String))
            dtUsers.Columns.Add("FullName", GetType(String))
            dtUsers.Columns.Add("Role", GetType(String))
            dtUsers.Columns.Add("Office", GetType(String))
            dtUsers.Columns.Add("CanRoute", GetType(Boolean))
            dtUsers.Columns.Add("CanMove", GetType(Boolean))
            dtUsers.Columns.Add("CanSoftCopy", GetType(Boolean))
            dtUsers.Columns.Add("IsActive", GetType(Boolean))
            dtUsers.Columns.Add("PendingSync", GetType(Boolean)).DefaultValue = False
            ' Auth matches a SHA-256 digest, never the raw credential: this XML file sits on
            ' a workstation disk. The mirror carries a masked display value and the hash;
            ' RFID_UID_RAW exists only so an unsynced offline enrolment can still replay its
            ' real card number to SQL, and is cleared once the row is mirrored.
            dtUsers.Columns.Add("RFID_UID_HASH", GetType(String))
            dtUsers.Columns.Add("RFID_UID_RAW", GetType(String))
            dtUsers.Columns.Add("IsLocked", GetType(Boolean)).DefaultValue = False
            dtUsers.PrimaryKey = New DataColumn() {dtUsers.Columns("UserID")}
            DataSet.Tables.Add(dtUsers)

            ' tbl_Documents
            Dim dtDocs As New DataTable("Documents")
            dtDocs.Columns.Add("DocumentID", GetType(Integer)).AutoIncrement = True
            dtDocs.Columns("DocumentID").AutoIncrementSeed = 1
            dtDocs.Columns("DocumentID").AutoIncrementStep = 1
            dtDocs.Columns.Add("DocCode", GetType(String))
            dtDocs.Columns.Add("DocType", GetType(String))
            dtDocs.Columns.Add("Title", GetType(String))
            dtDocs.Columns.Add("OriginatingOffice", GetType(String))
            dtDocs.Columns.Add("DestinationOffice", GetType(String))
            dtDocs.Columns.Add("CabinetID", GetType(String))
            dtDocs.Columns.Add("ShelfNo", GetType(String))
            dtDocs.Columns.Add("BoxCode", GetType(String))
            dtDocs.Columns.Add("GDriveURL", GetType(String))
            dtDocs.Columns.Add("CurrentStatus", GetType(String))
            dtDocs.Columns.Add("AssignedStaff", GetType(String))
            dtDocs.Columns.Add("DateReceived", GetType(String))
            dtDocs.Columns.Add("FlowDirection", GetType(String))
            dtDocs.Columns.Add("AssignedSection", GetType(String))
            dtDocs.Columns.Add("TargetDeadlineUTC", GetType(String))
            dtDocs.Columns.Add("RevisionPunchlist", GetType(String))
            dtDocs.Columns.Add("LastActionTaken", GetType(String))
            dtDocs.Columns.Add("ExternalControlNumber", GetType(String))
            dtDocs.Columns.Add("RequesterGender", GetType(String))
            dtDocs.Columns.Add("CreatedByUserID", GetType(Integer)).DefaultValue = 1
            dtDocs.Columns.Add("ModifiedByUserID", GetType(Integer)).DefaultValue = 1
            dtDocs.Columns.Add("IsNewOfflineRecord", GetType(Boolean)).DefaultValue = False
            dtDocs.Columns.Add("PendingSync", GetType(Boolean)).DefaultValue = False
            ' The SQL rowversion mirrored with each pull; the offline replay's optimistic
            ' guard compares it so a document another seat already moved is not overwritten.
            dtDocs.Columns.Add("RowVersion", GetType(Byte()))
            dtDocs.PrimaryKey = New DataColumn() {dtDocs.Columns("DocumentID")}
            DataSet.Tables.Add(dtDocs)

            ' tbl_ActionDirectives
            Dim dtDirectives As New DataTable("Directives")
            dtDirectives.Columns.Add("DirectiveID", GetType(Integer)).AutoIncrement = True
            dtDirectives.Columns("DirectiveID").AutoIncrementSeed = 1
            dtDirectives.Columns("DirectiveID").AutoIncrementStep = 1
            dtDirectives.Columns.Add("DocumentID", GetType(Integer))
            dtDirectives.Columns.Add("SGDirective", GetType(String))
            dtDirectives.Columns.Add("AssignedTo", GetType(String))
            dtDirectives.Columns.Add("Notes", GetType(String))
            dtDirectives.Columns.Add("LogUser", GetType(String))
            dtDirectives.Columns.Add("Timestamp", GetType(String))
            dtDirectives.Columns.Add("DirectiveCode", GetType(String))
            dtDirectives.Columns.Add("IssuedByUserID", GetType(Integer)).DefaultValue = 1
            dtDirectives.Columns.Add("PendingSync", GetType(Boolean)).DefaultValue = False
            dtDirectives.PrimaryKey = New DataColumn() {dtDirectives.Columns("DirectiveID")}
            DataSet.Tables.Add(dtDirectives)

            ' tbl_RoutingLogs
            Dim dtRouting As New DataTable("RoutingLogs")
            dtRouting.Columns.Add("RoutingID", GetType(Integer)).AutoIncrement = True
            dtRouting.Columns("RoutingID").AutoIncrementSeed = 1
            dtRouting.Columns("RoutingID").AutoIncrementStep = 1
            dtRouting.Columns.Add("DocumentID", GetType(Integer))
            dtRouting.Columns.Add("FromOffice", GetType(String))
            dtRouting.Columns.Add("ToOffice", GetType(String))
            dtRouting.Columns.Add("RoutedBy", GetType(String))
            dtRouting.Columns.Add("ActionTaken", GetType(String))
            dtRouting.Columns.Add("Remarks", GetType(String))
            dtRouting.Columns.Add("Timestamp", GetType(String))
            dtRouting.Columns.Add("RoutedByUserID", GetType(Integer)).DefaultValue = 1
            dtRouting.Columns.Add("PendingSync", GetType(Boolean)).DefaultValue = False
            dtRouting.PrimaryKey = New DataColumn() {dtRouting.Columns("RoutingID")}
            DataSet.Tables.Add(dtRouting)

            ' tbl_DocumentMovements
            Dim dtMovements As New DataTable("Movements")
            dtMovements.Columns.Add("MovementID", GetType(Integer)).AutoIncrement = True
            dtMovements.Columns("MovementID").AutoIncrementSeed = 1
            dtMovements.Columns("MovementID").AutoIncrementStep = 1
            dtMovements.Columns.Add("DocumentID", GetType(Integer))
            dtMovements.Columns.Add("FromLocation", GetType(String))
            dtMovements.Columns.Add("ToLocation", GetType(String))
            dtMovements.Columns.Add("MovedBy", GetType(String))
            dtMovements.Columns.Add("Reason", GetType(String))
            dtMovements.Columns.Add("Timestamp", GetType(String))
            dtMovements.Columns.Add("MovedByUserID", GetType(Integer)).DefaultValue = 1
            dtMovements.Columns.Add("PendingSync", GetType(Boolean)).DefaultValue = False
            dtMovements.PrimaryKey = New DataColumn() {dtMovements.Columns("MovementID")}
            DataSet.Tables.Add(dtMovements)

            ' tbl_AuditTrail
            Dim dtAudit As New DataTable("AuditTrail")
            dtAudit.Columns.Add("AuditID", GetType(Integer)).AutoIncrement = True
            dtAudit.Columns("AuditID").AutoIncrementSeed = 1
            dtAudit.Columns("AuditID").AutoIncrementStep = 1
            dtAudit.Columns.Add("UserName", GetType(String))
            dtAudit.Columns.Add("ActionDescription", GetType(String))
            dtAudit.Columns.Add("Timestamp", GetType(String))
            dtAudit.Columns.Add("ActionType", GetType(String))
            dtAudit.Columns.Add("UserID", GetType(Integer)).DefaultValue = 1
            dtAudit.Columns.Add("PendingSync", GetType(Boolean)).DefaultValue = False
            dtAudit.PrimaryKey = New DataColumn() {dtAudit.Columns("AuditID")}
            DataSet.Tables.Add(dtAudit)

            ' Workstation heartbeat mirror (tbl_WorkstationHeartbeat), read by the Admin
            ' tab's Workstation Sync Status grid. Keyed by machine name, not identity.
            Dim dtHeartbeat As New DataTable("Heartbeat")
            dtHeartbeat.Columns.Add("MachineName", GetType(String))
            dtHeartbeat.Columns.Add("LastSyncUTC", GetType(String))
            dtHeartbeat.Columns.Add("AppVersion", GetType(String))
            dtHeartbeat.Columns.Add("PendingOutbox", GetType(Integer))
            dtHeartbeat.Columns.Add("LastError", GetType(String))
            dtHeartbeat.Columns.Add("PendingSync", GetType(Boolean)).DefaultValue = False
            dtHeartbeat.PrimaryKey = New DataColumn() {dtHeartbeat.Columns("MachineName")}
            DataSet.Tables.Add(dtHeartbeat)
        End Sub

        Public Shared Sub AddUser(uid As String, name As String, role As String, Optional office As String = "", Optional canRoute As Boolean = True, Optional canMove As Boolean = True, Optional canSoftCopy As Boolean = True, Optional pendingSync As Boolean = False)
            Dim dt = DataSet.Tables("Users")
            Dim cleanUid = uid.Trim().ToUpperInvariant()
            Dim effOffice = If(String.IsNullOrWhiteSpace(office), role, office)
            Dim uidHash = If(cleanUid.Length > 0, Sha256Hex(cleanUid), "")
            Dim uidMask = If(cleanUid.Length > 0, MaskUid(cleanUid), "")
            Dim rawForReplay = If(pendingSync AndAlso cleanUid.Length > 0, cleanUid, Nothing)
            ' Match on any of the three carrier columns so rows from caches written before
            ' the hash column existed still resolve.
            Dim existing = dt.Select(String.Format("RFID_UID_HASH = '{0}' OR RFID_UID = '{1}' OR RFID_UID_RAW = '{1}'", uidHash.Replace("'", "''"), cleanUid.Replace("'", "''")))
            If existing.Length > 0 Then
                existing(0)("RFID_UID") = uidMask
                existing(0)("RFID_UID_HASH") = uidHash
                existing(0)("RFID_UID_RAW") = rawForReplay
                existing(0)("FullName") = name
                existing(0)("Role") = role
                If dt.Columns.Contains("Office") Then existing(0)("Office") = effOffice
                If dt.Columns.Contains("CanRoute") Then existing(0)("CanRoute") = canRoute
                If dt.Columns.Contains("CanMove") Then existing(0)("CanMove") = canMove
                If dt.Columns.Contains("CanSoftCopy") Then existing(0)("CanSoftCopy") = canSoftCopy
                existing(0)("IsActive") = True
                ' A pending row survives snapshot pulls; without the flag an offline
                ' enrolment or edit would be wiped by the next Users pull as unknown.
                If pendingSync AndAlso dt.Columns.Contains("PendingSync") Then existing(0)("PendingSync") = True
            Else
                Dim r = dt.NewRow()
                r("RFID_UID") = uidMask
                r("RFID_UID_HASH") = uidHash
                r("RFID_UID_RAW") = rawForReplay
                r("FullName") = name
                r("Role") = role
                If dt.Columns.Contains("Office") Then r("Office") = effOffice
                If dt.Columns.Contains("CanRoute") Then r("CanRoute") = canRoute
                If dt.Columns.Contains("CanMove") Then r("CanMove") = canMove
                If dt.Columns.Contains("CanSoftCopy") Then r("CanSoftCopy") = canSoftCopy
                r("IsActive") = True
                If dt.Columns.Contains("PendingSync") Then r("PendingSync") = pendingSync
                dt.Rows.Add(r)
            End If
            MarkDirty()
        End Sub

        Private Shared _terminalFailedTaps As Integer = 0
        Private Shared _terminalLockoutUntilUTC As DateTime = DateTime.MinValue

        Public Shared Function IsTerminalLockedOut() As Boolean
            Return DateTime.UtcNow < _terminalLockoutUntilUTC
        End Function

        Public Shared Sub ResetTerminalLockout()
            System.Threading.Interlocked.Exchange(_terminalFailedTaps, 0)
            _terminalLockoutUntilUTC = DateTime.MinValue
        End Sub

        ''' <summary>
        ''' The card credential is stored and matched only as a SHA-256 hex digest of the
        ''' uppercase ASCII card id. The Users mirror projection computes the same digest
        ''' with SQL Server's HASHBYTES('SHA2_256', CAST(... AS varchar)), so the two sides
        ''' hash the same bytes. Legacy caches that still carry the raw UID in RFID_UID
        ''' keep authenticating through the raw fallback branches.
        ''' </summary>
        Friend Shared Function Sha256Hex(value As String) As String
            Using sha = System.Security.Cryptography.SHA256.Create()
                Dim bytes = sha.ComputeHash(System.Text.Encoding.ASCII.GetBytes(If(value, "")))
                Dim sb As New System.Text.StringBuilder(bytes.Length * 2)
                For Each b In bytes
                    sb.Append(b.ToString("X2"))
                Next
                Return sb.ToString()
            End Using
        End Function

        Private Shared Function MaskUid(uid As String) As String
            Return If(If(uid, "").Length > 4, "****" & uid.Substring(uid.Length - 4), uid)
        End Function

        ' The offline terminal lockout reads the same configured threshold the connected
        ' path enforces at FormMain, so a tuned setting is not silently ignored offline.
        Private Shared ReadOnly Property TerminalLockoutThreshold As Integer
            Get
                Dim threshold As Integer = 5
                Try
                    threshold = AppSettings.Instance.RfidSettings.LockoutThreshold
                Catch
                End Try
                Return If(threshold > 0, threshold, 5)
            End Get
        End Property

        Public Shared Function AuthenticateRFID(uid As String) As DataRow
            EnsureInitialized()
            If IsTerminalLockedOut() Then
                Return Nothing
            End If

            If String.IsNullOrWhiteSpace(uid) Then Return Nothing
            Dim cleanUid = uid.Trim().ToUpperInvariant()
            ' The RFID listener thread and the UI can tap concurrently while the snapshot
            ' poller mutates the same tables, so the lookup runs under the store lock.
            SyncLock _syncLock
                Dim rows = DataSet.Tables("Users").Select(String.Format(
                    "(RFID_UID_HASH = '{0}' OR RFID_UID = '{1}' OR RFID_UID_RAW = '{1}') AND IsActive = True",
                    Sha256Hex(cleanUid).Replace("'", "''"), cleanUid.Replace("'", "''")))
                If rows.Length > 0 Then
                    System.Threading.Interlocked.Exchange(_terminalFailedTaps, 0)
                    Return rows(0)
                End If
            End SyncLock
            If System.Threading.Interlocked.Increment(_terminalFailedTaps) >= TerminalLockoutThreshold Then
                _terminalLockoutUntilUTC = DateTime.UtcNow.AddMinutes(5)
            End If
            Return Nothing
        End Function

        Public Shared Function GenerateDocCode(docType As String) As String
            EnsureInitialized()
            Dim prefix As String = "DOC"
            Select Case docType.Trim().ToLower()
                Case "regular communication", "comm", "regular_comm", "reg_comm" : prefix = "COMM"
                Case "legislative", "leg" : prefix = "LEG"
                Case "finance", "fin" : prefix = "FIN"
                Case "travel order", "travel", "to" : prefix = "TO"
                Case "resolution" : prefix = "RES"
                Case "parliament bill" : prefix = "BLL"
                Case "committee report" : prefix = "REP"
                Case "executive communication" : prefix = "EXC"
                Case "memorandum" : prefix = "MEM"
                Case "endorsement" : prefix = "END"
                Case "journal entry" : prefix = "JRN"
            End Select

            Dim yearStr As String = DateTime.Now.Year.ToString()
            Dim matchingRows = DataSet.Tables("Documents").Select(String.Format("DocCode LIKE '{0}-{1}-%'", prefix, yearStr))
            Dim nextSeq As Integer = matchingRows.Length + 1
            Return String.Format("{0}-{1}-{2:D3}", prefix, yearStr, nextSeq)
        End Function

        Public Shared Function ValidateGDriveURL(url As String, ByRef errMessage As String) As Boolean
            errMessage = ""
            If String.IsNullOrWhiteSpace(url) Then Return True
            Dim trimmed = url.Trim()
            If System.IO.File.Exists(trimmed) Then
                Dim ext = System.IO.Path.GetExtension(trimmed).ToLowerInvariant()
                If ext <> ".pdf" Then
                    errMessage = "Security restriction: Only PDF files (.pdf) may be linked or opened as scans."
                    Return False
                End If
                Return True
            End If
            If Not trimmed.StartsWith("https://", StringComparison.OrdinalIgnoreCase) Then
                errMessage = "URL must use secure 'https://' protocol or reference an accessible local .pdf file."
                Return False
            End If
            Try
                Dim uri As New Uri(trimmed)
                Dim host = uri.Host.ToLowerInvariant()
                If host <> "drive.google.com" AndAlso host <> "docs.google.com" AndAlso Not host.EndsWith(".google.com") Then
                    errMessage = "Allowed domain hosts are restricted to 'drive.google.com' or 'docs.google.com'."
                    Return False
                End If
                Return True
            Catch ex As Exception
                errMessage = "Malformed URL format."
                Return False
            End Try
        End Function

        Public Shared Function AddDocument(code As String, docType As String, title As String, origin As String, dest As String, cab As String, shelf As String, box As String, url As String, status As String, assigned As String, Optional flowDirection As String = "INCOMING", Optional assignedSection As String = "", Optional targetDeadline As String = "", Optional punchlist As String = "", Optional lastAction As String = "", Optional externalControlNumber As String = "", Optional isOffline As Boolean = True, Optional createdByUserId As Integer = 1, Optional requesterGender As String = "") As Integer
            SyncLock _syncLock
                EnsureInitialized()
                Dim dt = DataSet.Tables("Documents")
                Dim r = dt.NewRow()
                r("DocCode") = code
                r("DocType") = docType
                r("Title") = title
                r("OriginatingOffice") = origin
                r("DestinationOffice") = dest
                r("CabinetID") = cab
                r("ShelfNo") = shelf
                r("BoxCode") = box
                r("GDriveURL") = url
                r("CurrentStatus") = status
                r("AssignedStaff") = assigned
                r("DateReceived") = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                r("FlowDirection") = If(String.IsNullOrEmpty(flowDirection), "INCOMING", flowDirection)
                r("AssignedSection") = If(String.IsNullOrEmpty(assignedSection), DocumentService.GetDefaultSectionForCategory(docType), assignedSection)
                r("TargetDeadlineUTC") = targetDeadline
                r("RevisionPunchlist") = punchlist
                r("LastActionTaken") = If(String.IsNullOrEmpty(lastAction), "Registered and routed to " & r("AssignedSection").ToString(), lastAction)
                r("ExternalControlNumber") = If(externalControlNumber, "")
                If dt.Columns.Contains("RequesterGender") Then r("RequesterGender") = If(requesterGender, "")
                If dt.Columns.Contains("CreatedByUserID") Then r("CreatedByUserID") = createdByUserId
                If dt.Columns.Contains("ModifiedByUserID") Then r("ModifiedByUserID") = createdByUserId
                If dt.Columns.Contains("IsNewOfflineRecord") Then r("IsNewOfflineRecord") = isOffline
                If dt.Columns.Contains("PendingSync") Then r("PendingSync") = isOffline
                dt.Rows.Add(r)
                MarkDirty()

                Dim docId = CInt(r("DocumentID"))
                AddRoutingLog(docId, origin, dest, assigned, "REGISTERED", "Document registered in OSG System.", isOffline, createdByUserId)
                AddMovementLog(docId, "INCOMING", String.Format("{0}/{1}/{2}", cab, shelf, box), assigned, "Initial storage assignment.", isOffline, createdByUserId)
                Return docId
            End SyncLock
        End Function

        Public Shared Sub AddDirective(docId As Integer, directive As String, assignedTo As String, notes As String, staffName As String, Optional isOffline As Boolean = True, Optional issuedByUserId As Integer = 1, Optional directiveCode As String = "IMMEDIATE_ACTION")
            SyncLock _syncLock
                Dim dt = DataSet.Tables("Directives")
                Dim r = dt.NewRow()
                r("DocumentID") = docId
                r("SGDirective") = directive
                r("AssignedTo") = assignedTo
                r("Notes") = notes
                r("LogUser") = staffName
                r("Timestamp") = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                If dt.Columns.Contains("DirectiveCode") Then r("DirectiveCode") = directiveCode
                If dt.Columns.Contains("IssuedByUserID") Then r("IssuedByUserID") = issuedByUserId
                If dt.Columns.Contains("PendingSync") Then r("PendingSync") = isOffline
                dt.Rows.Add(r)

                Dim docRows = DataSet.Tables("Documents").Select("DocumentID = " & docId)
                If docRows.Length > 0 Then
                    ' Only directives with a real status outcome move the status column; the
                    ' directive text itself belongs in LastActionTaken. Free text in
                    ' CurrentStatus replayed to a RECEIVED fallback corrupted the status.
                    ' "Approved & Archived" maps to ARCHIVED to match the connected path,
                    ' where the APPROVE_ARCHIVE type's ResultStatusID is the archived status.
                    If directive = "REVISION_REQUESTED" Then
                        docRows(0)("CurrentStatus") = "FOR_REVISION"
                    ElseIf directive = "APPROVED" Then
                        docRows(0)("CurrentStatus") = "APPROVED"
                    ElseIf directive = "Approved & Archived" Then
                        docRows(0)("CurrentStatus") = "ARCHIVED"
                    End If
                    docRows(0)("LastActionTaken") = "SG Directive: " & directive
                    If Not String.IsNullOrEmpty(assignedTo) Then docRows(0)("AssignedStaff") = assignedTo
                    If docRows(0).Table.Columns.Contains("ModifiedByUserID") Then docRows(0)("ModifiedByUserID") = issuedByUserId
                    If isOffline AndAlso docRows(0).Table.Columns.Contains("PendingSync") Then
                        docRows(0)("PendingSync") = True
                    End If
                End If
                MarkDirty()
            End SyncLock
        End Sub

        Public Shared Sub AddRoutingLog(docId As Integer, fromOffice As String, toOffice As String, routedBy As String, action As String, remarks As String, Optional isOffline As Boolean = True, Optional routedByUserId As Integer = 1)
            SyncLock _syncLock
                Dim dt = DataSet.Tables("RoutingLogs")
                Dim r = dt.NewRow()
                r("DocumentID") = docId
                r("FromOffice") = fromOffice
                r("ToOffice") = toOffice
                r("RoutedBy") = routedBy
                r("ActionTaken") = action
                r("Remarks") = remarks
                r("Timestamp") = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                If dt.Columns.Contains("RoutedByUserID") Then r("RoutedByUserID") = routedByUserId
                If dt.Columns.Contains("PendingSync") Then r("PendingSync") = isOffline
                dt.Rows.Add(r)

                Dim docRows = DataSet.Tables("Documents").Select("DocumentID = " & docId)
                If docRows.Length > 0 Then
                    docRows(0)("OriginatingOffice") = fromOffice
                    docRows(0)("DestinationOffice") = toOffice
                    If Not action.Equals("REVISION_REQUESTED", StringComparison.OrdinalIgnoreCase) Then
                        docRows(0)("LastActionTaken") = "Routed to " & toOffice & ": " & action
                    End If
                    If docRows(0).Table.Columns.Contains("ModifiedByUserID") Then docRows(0)("ModifiedByUserID") = routedByUserId
                    If isOffline AndAlso docRows(0).Table.Columns.Contains("PendingSync") Then
                        docRows(0)("PendingSync") = True
                    End If
                End If
                MarkDirty()
            End SyncLock
        End Sub

        Public Shared Sub AddMovementLog(docId As Integer, fromLoc As String, toLoc As String, movedBy As String, reason As String, Optional isOffline As Boolean = True, Optional movedByUserId As Integer = 1)
            SyncLock _syncLock
                Dim dt = DataSet.Tables("Movements")
                Dim r = dt.NewRow()
                r("DocumentID") = docId
                r("FromLocation") = fromLoc
                r("ToLocation") = toLoc
                r("MovedBy") = movedBy
                r("Reason") = reason
                r("Timestamp") = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                If dt.Columns.Contains("MovedByUserID") Then r("MovedByUserID") = movedByUserId
                If dt.Columns.Contains("PendingSync") Then r("PendingSync") = isOffline
                dt.Rows.Add(r)

                Dim parts = toLoc.Split(New Char() {"/"c, "|"c})
                Dim docRows = DataSet.Tables("Documents").Select("DocumentID = " & docId)
                If docRows.Length > 0 Then
                    If parts.Length >= 3 Then
                        docRows(0)("CabinetID") = parts(0).Trim()
                        docRows(0)("ShelfNo") = parts(1).Trim()
                        docRows(0)("BoxCode") = parts(2).Trim()
                    End If
                    docRows(0)("LastActionTaken") = "Moved storage to " & toLoc
                    If docRows(0).Table.Columns.Contains("ModifiedByUserID") Then docRows(0)("ModifiedByUserID") = movedByUserId
                    If isOffline AndAlso docRows(0).Table.Columns.Contains("PendingSync") Then
                        docRows(0)("PendingSync") = True
                    End If
                End If
                MarkDirty()
            End SyncLock
        End Sub

        Public Shared Sub LogAudit(user As String, action As String, Optional isOffline As Boolean = True, Optional userId As Integer = 1, Optional actionType As String = "")
            SyncLock _syncLock
                Dim dt = DataSet.Tables("AuditTrail")
                Dim r = dt.NewRow()
                r("UserName") = user
                r("ActionDescription") = action
                r("Timestamp") = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                If dt.Columns.Contains("ActionType") Then r("ActionType") = actionType
                If dt.Columns.Contains("UserID") Then r("UserID") = userId
                If dt.Columns.Contains("PendingSync") Then r("PendingSync") = isOffline
                dt.Rows.Add(r)
                MarkDirty()
            End SyncLock
        End Sub

        Public Shared Sub RequestRevision(docId As Integer, punchlistNotes As String, returnSection As String, staffName As String, Optional staffUserId As Integer = 1)
            Dim extCn As String = ""
            SyncLock _syncLock
                Dim doc = FindDocumentRow(docId)
                If doc Is Nothing Then Return

                doc("CurrentStatus") = "FOR_REVISION"
                Dim prevPunchlist As String = doc("RevisionPunchlist").ToString()
                Dim stamp As String = "[" & DateTime.Now.ToString("yyyy-MM-dd HH:mm") & "] " & punchlistNotes
                doc("RevisionPunchlist") = If(String.IsNullOrWhiteSpace(prevPunchlist), stamp, prevPunchlist & vbCrLf & stamp)

                Dim targetSec As String = If(Not String.IsNullOrWhiteSpace(returnSection), returnSection, DocumentService.GetDefaultSectionForCategory(doc("DocType").ToString()))
                doc("AssignedSection") = targetSec
                doc("LastActionTaken") = "Revision requested by " & staffName & ": " & punchlistNotes
                If doc.Table.Columns.Contains("ModifiedByUserID") Then doc("ModifiedByUserID") = staffUserId
                MarkDirty()

                AddDirective(docId, "REVISION_REQUESTED", targetSec, punchlistNotes, staffName, isOffline:=True, issuedByUserId:=staffUserId)
                AddRoutingLog(docId, "Office of the Secretary-General", targetSec, staffName, "REVISION_REQUESTED", punchlistNotes, isOffline:=True, routedByUserId:=staffUserId)
                LogAudit(staffName, "Revision requested for Document #" & docId & " with punchlist: " & punchlistNotes, isOffline:=True, userId:=staffUserId)

                extCn = RowString(doc, "ExternalControlNumber")
            End SyncLock

            PushPortalStatusFor(extCn, "FOR_REVISION")
        End Sub

        Public Shared Sub ResubmitDocument(docId As Integer, staffName As String, notes As String, Optional staffUserId As Integer = 1)
            Dim extCn As String = ""
            SyncLock _syncLock
                Dim doc = FindDocumentRow(docId)
                If doc Is Nothing Then Return

                Dim originSec As String = doc("AssignedSection").ToString()
                doc("CurrentStatus") = "FOR_REVIEW"
                doc("AssignedSection") = "Secretary-General"
                doc("LastActionTaken") = "Resubmitted for Sec Gen review: " & notes
                If doc.Table.Columns.Contains("ModifiedByUserID") Then doc("ModifiedByUserID") = staffUserId
                MarkDirty()

                AddRoutingLog(docId, originSec, "Office of the Secretary-General", staffName, "RESUBMITTED", notes, isOffline:=True, routedByUserId:=staffUserId)
                LogAudit(staffName, "Document #" & docId & " resubmitted by " & originSec & " to Sec Gen: " & notes, isOffline:=True, userId:=staffUserId)

                extCn = RowString(doc, "ExternalControlNumber")
            End SyncLock

            PushPortalStatusFor(extCn, "FOR_REVIEW")
        End Sub

        Public Shared Sub ApproveDocument(docId As Integer, staffName As String, notes As String, Optional staffUserId As Integer = 1)
            Dim extCn As String = ""
            SyncLock _syncLock
                Dim doc = FindDocumentRow(docId)
                If doc Is Nothing Then Return

                doc("CurrentStatus") = "APPROVED"
                doc("AssignedSection") = "Records Section"
                doc("LastActionTaken") = "Approved by " & staffName & ": " & notes
                If doc.Table.Columns.Contains("ModifiedByUserID") Then doc("ModifiedByUserID") = staffUserId
                MarkDirty()

                AddDirective(docId, "APPROVED", "Records Section", notes, staffName, isOffline:=True, issuedByUserId:=staffUserId)
                AddRoutingLog(docId, "Office of the Secretary-General", "Records Section", staffName, "APPROVED", notes, isOffline:=True, routedByUserId:=staffUserId)
                LogAudit(staffName, "Document #" & docId & " approved by " & staffName & ": " & notes, isOffline:=True, userId:=staffUserId)

                extCn = RowString(doc, "ExternalControlNumber")
            End SyncLock

            PushPortalStatusFor(extCn, "APPROVED")
        End Sub

        Public Shared Sub ReleaseDocument(docId As Integer, staffName As String, notes As String, Optional staffUserId As Integer = 1)
            Dim extCn As String = ""
            SyncLock _syncLock
                Dim doc = FindDocumentRow(docId)
                If doc Is Nothing Then Return

                Dim dest As String = doc("DestinationOffice").ToString()
                doc("CurrentStatus") = "RELEASED"
                doc("AssignedSection") = "Archived / Released"
                doc("LastActionTaken") = "Released to " & dest & ": " & notes
                If doc.Table.Columns.Contains("ModifiedByUserID") Then doc("ModifiedByUserID") = staffUserId
                MarkDirty()

                AddRoutingLog(docId, "Records Section", dest, staffName, "RELEASED", notes, isOffline:=True, routedByUserId:=staffUserId)
                LogAudit(staffName, "Document #" & docId & " released: " & notes, isOffline:=True, userId:=staffUserId)

                extCn = RowString(doc, "ExternalControlNumber")
            End SyncLock

            PushPortalStatusFor(extCn, "RELEASED")
        End Sub

        Private Shared Function FindDocumentRow(docId As Integer) As DataRow
            EnsureInitialized()
            Dim dt = DataSet.Tables("Documents")
            If dt Is Nothing Then Return Nothing
            Dim rows = dt.Select("DocumentID = " & docId)
            If rows.Length = 0 Then Return Nothing
            Return rows(0)
        End Function

        ' The portal push runs after _syncLock releases: it performs network I/O and must not
        ' hold the cache lock across it.
        Private Shared Sub PushPortalStatusFor(externalControlNumber As String, statusCode As String)
            If Not String.IsNullOrWhiteSpace(externalControlNumber) Then
                DocumentService.PushPortalStatusSafe(externalControlNumber, statusCode)
            End If
        End Sub

        ''' <summary>
        ''' Returns a live DataView over the Documents table instead of a copied table, so
        ''' refresh-heavy screens bind without an O(rows x columns) deep copy per call.
        ''' </summary>
        Public Shared Function GetVisibleDocuments(user As DataRow, Optional sectionFilter As String = "", Optional categoryFilter As String = "") As DataView
            SyncLock _syncLock
                EnsureInitialized()
                Dim dt = DataSet.Tables("Documents")
                ' No session, no documents: a logged-out desk must not browse the registry,
                ' and section isolation has no anonymous case.
                If user Is Nothing Then Return New DataView(dt.Clone())

                Dim role As String = user("Role").ToString()
                Dim office As String = If(user.Table.Columns.Contains("Office") AndAlso Not IsDBNull(user("Office")), user("Office").ToString(), "")
                Dim effectiveDesk As String = If(Not String.IsNullOrWhiteSpace(office), office, role)
                Dim isManager As Boolean = (role = "Secretary-General" OrElse role = "System Administrator" OrElse role = "OSG Chief")
                Dim effectiveSection As String = sectionFilter.Trim()

                If Not isManager AndAlso String.IsNullOrEmpty(effectiveSection) Then
                    If effectiveDesk.EndsWith("Section", StringComparison.OrdinalIgnoreCase) OrElse effectiveDesk.Equals("Secretariat", StringComparison.OrdinalIgnoreCase) Then
                        effectiveSection = effectiveDesk
                    End If
                End If

                Dim filters As New List(Of String)()

                If Not String.IsNullOrEmpty(effectiveSection) AndAlso Not effectiveSection.Equals("All Sections", StringComparison.OrdinalIgnoreCase) Then
                    filters.Add(String.Format("AssignedSection = '{0}'", effectiveSection.Replace("'", "''")))
                ElseIf Not isManager Then
                    Dim staffName As String = user("FullName").ToString().Replace("'", "''")
                    filters.Add(String.Format("(AssignedStaff = '{0}' OR AssignedStaff = '' OR AssignedStaff IS NULL)", staffName))
                End If

                If Not String.IsNullOrEmpty(categoryFilter) AndAlso Not categoryFilter.Equals("All Categories", StringComparison.OrdinalIgnoreCase) AndAlso Not categoryFilter.Equals("All", StringComparison.OrdinalIgnoreCase) Then
                    filters.Add(String.Format("DocType = '{0}'", categoryFilter.Replace("'", "''")))
                End If

                Dim dv As New DataView(dt)
                If filters.Count > 0 Then
                    dv.RowFilter = String.Join(" AND ", filters)
                End If
                Return dv
            End SyncLock
        End Function

        Public Shared Function MapRowToDocument(row As DataRow) As Document
            If row Is Nothing Then Return Nothing
            Dim docObj As New Document With {
                .DocumentID = RowInt(row, "DocumentID"),
                .DocCode = RowString(row, "DocCode"),
                .Title = RowString(row, "Title"),
                .OriginOffice = RowString(row, "OriginatingOffice"),
                .DestinationOffice = RowString(row, "DestinationOffice"),
                .FlowDirection = RowString(row, "FlowDirection", "INCOMING"),
                .AssignedSection = RowString(row, "AssignedSection"),
                .RevisionPunchlist = RowString(row, "RevisionPunchlist"),
                .LastActionTaken = RowString(row, "LastActionTaken"),
                .RegisteredAtUTC = DateTime.UtcNow
            }

            docObj.Remarks = RowStringOrNull(row, "Remarks")

            Dim typeStr = RowString(row, "DocType").ToUpperInvariant()
            Select Case typeStr
                Case "FINANCE", "FIN": docObj.DocumentTypeID = 3
                Case "TRAVEL", "TRAVEL ORDER", "TO": docObj.DocumentTypeID = 4
                Case "LEGISLATIVE", "LEG": docObj.DocumentTypeID = 2
                Case Else: docObj.DocumentTypeID = 1
            End Select

            Dim statusStr = RowString(row, "CurrentStatus").ToUpperInvariant()
            Select Case statusStr
                Case "RECEIVED", "LOGGED": docObj.StatusID = 1
                Case "FOR_REVIEW": docObj.StatusID = 2
                Case "FOR_REVISION": docObj.StatusID = 3
                Case "APPROVED": docObj.StatusID = 6
                Case "RELEASED": docObj.StatusID = 7
                Case "FILED", "ARCHIVED": docObj.StatusID = 8
                Case Else: docObj.StatusID = 1
            End Select

            docObj.TargetDeadlineUTC = RowDateTimeOrNull(row, "TargetDeadlineUTC")
            docObj.CurrentStorageLocationID = RowIntOrNull(row, "CurrentStorageLocationID")

            docObj.CabinetID = RowStringOrNull(row, "CabinetID")
            docObj.ShelfNo = RowStringOrNull(row, "ShelfNo")
            docObj.BoxCode = RowStringOrNull(row, "BoxCode")
            docObj.ExternalControlNumber = RowStringOrNull(row, "ExternalControlNumber")

            Return docObj
        End Function

        ' Cache rows come from two writers (embedded XML and SQL mirrors) at different schema
        ' generations, so every column read must tolerate the column being absent entirely.
        Private Shared Function RowStringOrNull(row As DataRow, columnName As String) As String
            If Not row.Table.Columns.Contains(columnName) OrElse IsDBNull(row(columnName)) Then Return Nothing
            Return row(columnName).ToString()
        End Function

        Private Shared Function RowString(row As DataRow, columnName As String, Optional fallback As String = "") As String
            Return If(RowStringOrNull(row, columnName), fallback)
        End Function

        Private Shared Function RowIntOrNull(row As DataRow, columnName As String) As Integer?
            If Not row.Table.Columns.Contains(columnName) OrElse IsDBNull(row(columnName)) Then Return Nothing
            Dim parsed As Integer
            If Integer.TryParse(row(columnName).ToString(), parsed) Then Return parsed
            Return Nothing
        End Function

        Private Shared Function RowInt(row As DataRow, columnName As String) As Integer
            Dim parsed = RowIntOrNull(row, columnName)
            Return If(parsed.HasValue, parsed.Value, 0)
        End Function

        Private Shared Function RowDateTimeOrNull(row As DataRow, columnName As String) As DateTime?
            If Not row.Table.Columns.Contains(columnName) OrElse IsDBNull(row(columnName)) Then Return Nothing
            Dim parsed As DateTime
            If DateTime.TryParse(row(columnName).ToString(), parsed) Then Return parsed
            Return Nothing
        End Function

        Public Shared Function GetDocumentByID(docId As Integer) As Document
            SyncLock _syncLock
                Dim doc = FindDocumentRow(docId)
                If doc Is Nothing Then Return Nothing
                Return MapRowToDocument(doc)
            End SyncLock
        End Function

        Public Shared Function GetRoutingLogsForDocument(docId As Integer) As List(Of RoutingLog)
            SyncLock _syncLock
                EnsureInitialized()
                Dim list As New List(Of RoutingLog)()
                Dim dt = DataSet.Tables("RoutingLogs")
                If dt Is Nothing Then Return list

                Dim rows = dt.Select("DocumentID = " & docId, "RoutingID ASC")
                For Each r In rows
                    Dim logEntry As New RoutingLog With {
                        .RoutingLogID = Convert.ToInt32(r("RoutingID")),
                        .DocumentID = docId,
                        .FromOffice = r("FromOffice").ToString(),
                        .ToOffice = r("ToOffice").ToString(),
                        .RoutingRemarks = r("Remarks").ToString(),
                        .RoutedAtUTC = DateTime.UtcNow
                    }
                    Dim routedAt = RowDateTimeOrNull(r, "Timestamp")
                    If routedAt.HasValue Then logEntry.RoutedAtUTC = routedAt.Value
                    list.Add(logEntry)
                Next
                Return list
            End SyncLock
        End Function

        Public Shared Function GetDirectivesForDocument(docId As Integer) As List(Of ActionDirective)
            SyncLock _syncLock
                EnsureInitialized()
                Dim list As New List(Of ActionDirective)()
                Dim dt = DataSet.Tables("Directives")
                If dt Is Nothing Then Return list

                Dim rows = dt.Select("DocumentID = " & docId, "DirectiveID ASC")
                For Each r In rows
                    Dim dirEntry As New ActionDirective With {
                        .DirectiveID = Convert.ToInt32(r("DirectiveID")),
                        .DocumentID = docId,
                        .DirectiveText = r("SGDirective").ToString() & " : " & r("Notes").ToString(),
                        .Remarks = r("Notes").ToString(),
                        .IssuedAtUTC = DateTime.UtcNow
                    }
                    Dim issuedAt = RowDateTimeOrNull(r, "Timestamp")
                    If issuedAt.HasValue Then dirEntry.IssuedAtUTC = issuedAt.Value
                    list.Add(dirEntry)
                Next
                Return list
            End SyncLock
        End Function
    End Class
End Namespace
