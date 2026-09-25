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
        Private Shared ReadOnly DbPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "bta_osg_db.xml")
        Public Shared DataSet As New DataSet("BTA_OSG_DB")

        Public Shared Sub Initialize()
            DataSet.Clear()
            DataSet.Tables.Clear()

            CreateTables()

            If File.Exists(DbPath) Then
                Try
                    DataSet.ReadXml(DbPath)
                    If DataSet.Tables.Contains("Documents") AndAlso Not DataSet.Tables("Documents").Columns.Contains("ExternalControlNumber") Then
                        DataSet.Tables("Documents").Columns.Add("ExternalControlNumber", GetType(String))
                    End If
                    For Each tblName In {"Users", "Documents", "Directives", "RoutingLogs", "Movements", "AuditTrail"}
                        If DataSet.Tables.Contains(tblName) AndAlso Not DataSet.Tables(tblName).Columns.Contains("PendingSync") Then
                            DataSet.Tables(tblName).Columns.Add("PendingSync", GetType(Boolean)).DefaultValue = False
                        End If
                    Next
                    EnsureDefaultUsers()
                    Return
                Catch ex As Exception
                    ' Fallback to re-creating if corrupt
                End Try
            End If

            SeedDefaultData()
            Save()
        End Sub

        Public Shared Sub EnsureInitialized()
            If DataSet.Tables.Count = 0 Then
                Initialize()
            End If
        End Sub

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
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine("EmbeddedDB.Save failed: " & ex.Message)
                End Try
            End SyncLock
        End Sub

        ''' <summary>
        ''' Merges one SQL Server snapshot into its cache table. SQL Server is the source of
        ''' truth whenever it is reachable. Rows are updated in place rather than cleared and reloaded.
        ''' Rows marked with PendingSync = True are strictly preserved so offline work is never wiped.
        ''' </summary>
        Public Shared Sub ApplySnapshot(tableName As String, snapshot As DataTable, Optional persist As Boolean = True)
            If snapshot Is Nothing Then Return
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

                ' Apply the snapshot first, then drop what it no longer lists (unless pending sync)
                Using reader = snapshot.CreateDataReader()
                    target.Load(reader, LoadOption.OverwriteChanges)
                End Using

                Dim doomed As New List(Of DataRow)()
                For Each row As DataRow In target.Rows
                    Dim isPending = target.Columns.Contains("PendingSync") AndAlso Not IsDBNull(row("PendingSync")) AndAlso CBool(row("PendingSync"))
                    If Not isPending AndAlso Not incoming.Contains(row(pk)) Then
                        doomed.Add(row)
                    End If
                Next
                For Each row As DataRow In doomed
                    target.Rows.Remove(row)
                Next

                If persist Then MarkDirty()
            End SyncLock
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
            dtDocs.Columns.Add("PendingSync", GetType(Boolean)).DefaultValue = False
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
            dtAudit.Columns.Add("PendingSync", GetType(Boolean)).DefaultValue = False
            dtAudit.PrimaryKey = New DataColumn() {dtAudit.Columns("AuditID")}
            DataSet.Tables.Add(dtAudit)
        End Sub

        Private Shared Sub EnsureDefaultUsers()
            AddUser("88A9F321", "Prof. Ali B. Pangalian", "Secretary-General", "Office of the Secretary-General")
            AddUser("99B1C456", "Atty. Fatima Z. Rasheed", "Legislative Section", "Legislative Section")
            AddUser("77C3D987", "Omire Khalid B. Ebrahim", "System Administrator", "ICT / Systems Administration")
            AddUser("55E5F666", "Hassim A. Ibrahim", "Finance Section", "Finance Section")
            AddUser("11A2B3C4", "CJ Fairoz A. Usop", "Travel Section", "Travel Section")
            AddUser("66A1B2C3", "Sittie K. Amin", "Records Section", "Records Section")
            AddUser("44C5D6E7", "Amina T. Macacua", "Secretariat", "Secretariat")
        End Sub

        Private Shared Sub SeedDefaultData()
            EnsureDefaultUsers()

            Dim id1 = AddDocument("COMM-2026-001", "Regular Communication", "Transmittal of Inter-Agency Cooperation Agreement", "Ministry of Interior", "Office of the Secretary-General", "CAB-A", "S-1", "BOX-01", "https://drive.google.com/file/d/sample-comm-001/view", "FOR_REVIEW", "Amina T. Macacua", "INCOMING", "Secretariat", DateTime.Now.AddDays(3).ToString("yyyy-MM-dd HH:mm:ss"), "", "Routed to Secretariat for initial review", "", isOffline:=False)
            AddRoutingLog(id1, "Records Section", "Secretariat", "Sittie K. Amin", "INTAKE_ROUTED", "Initial categorization completed.", isOffline:=False)

            Dim id2 = AddDocument("LEG-2026-001", "Legislative", "Bangsamoro Education Code Amendment Bill of 2026", "Committee on Education", "Legislative Section", "CAB-B", "S-3", "BOX-04", "https://drive.google.com/file/d/sample-leg-001/view", "FOR_REVIEW", "Atty. Fatima Z. Rasheed", "INCOMING", "Legislative Section", DateTime.Now.AddDays(7).ToString("yyyy-MM-dd HH:mm:ss"), "", "Referred to Legislative Section for committee report drafting", "", isOffline:=False)
            AddRoutingLog(id2, "Records Section", "Legislative Section", "Sittie K. Amin", "INTAKE_ROUTED", "Legislative categorization completed.", isOffline:=False)

            Dim id3 = AddDocument("FIN-2026-001", "Finance", "Q1 Parliament Operations Operating Budget Allocation", "Finance Division", "Finance Section", "CAB-C", "S-2", "BOX-02", "https://drive.google.com/file/d/sample-fin-001/view", "RECEIVED", "Hassim A. Ibrahim", "INCOMING", "Finance Section", DateTime.Now.AddDays(2).ToString("yyyy-MM-dd HH:mm:ss"), "", "Awaiting financial compliance verification", "", isOffline:=False)
            AddRoutingLog(id3, "Records Section", "Finance Section", "Sittie K. Amin", "INTAKE_ROUTED", "Finance intake registered.", isOffline:=False)

            Dim id4 = AddDocument("TO-2026-001", "Travel Order", "Official Mission Order to Davao City Consultation", "OSG Travel Desk", "Travel Section", "CAB-A", "S-2", "BOX-03", "https://drive.google.com/file/d/sample-to-001/view", "FOR_REVIEW", "CJ Fairoz A. Usop", "INCOMING", "Travel Section", DateTime.Now.AddDays(1).ToString("yyyy-MM-dd HH:mm:ss"), "", "Urgent travel authority review", "", isOffline:=False)
            AddRoutingLog(id4, "Records Section", "Travel Section", "Sittie K. Amin", "INTAKE_ROUTED", "Travel order registered.", isOffline:=False)

            LogAudit("SYSTEM", "Embedded Database Engine initialized with target OSG seed dataset.", isOffline:=False)
        End Sub

        Public Shared Sub AddUser(uid As String, name As String, role As String, Optional office As String = "", Optional canRoute As Boolean = True, Optional canMove As Boolean = True, Optional canSoftCopy As Boolean = True)
            Dim dt = DataSet.Tables("Users")
            Dim cleanUid = uid.Trim().ToUpperInvariant()
            Dim effOffice = If(String.IsNullOrWhiteSpace(office), role, office)
            Dim existing = dt.Select(String.Format("RFID_UID = '{0}'", cleanUid.Replace("'", "''")))
            If existing.Length > 0 Then
                existing(0)("FullName") = name
                existing(0)("Role") = role
                If dt.Columns.Contains("Office") Then existing(0)("Office") = effOffice
                If dt.Columns.Contains("CanRoute") Then existing(0)("CanRoute") = canRoute
                If dt.Columns.Contains("CanMove") Then existing(0)("CanMove") = canMove
                If dt.Columns.Contains("CanSoftCopy") Then existing(0)("CanSoftCopy") = canSoftCopy
                existing(0)("IsActive") = True
            Else
                Dim r = dt.NewRow()
                r("RFID_UID") = cleanUid
                r("FullName") = name
                r("Role") = role
                If dt.Columns.Contains("Office") Then r("Office") = effOffice
                If dt.Columns.Contains("CanRoute") Then r("CanRoute") = canRoute
                If dt.Columns.Contains("CanMove") Then r("CanMove") = canMove
                If dt.Columns.Contains("CanSoftCopy") Then r("CanSoftCopy") = canSoftCopy
                r("IsActive") = True
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
            _terminalFailedTaps = 0
            _terminalLockoutUntilUTC = DateTime.MinValue
        End Sub

        Public Shared Function AuthenticateRFID(uid As String) As DataRow
            EnsureInitialized()
            If IsTerminalLockedOut() Then
                Return Nothing
            End If

            If String.IsNullOrWhiteSpace(uid) Then Return Nothing
            Dim cleanUid = uid.Trim().ToUpperInvariant()
            Dim rows = DataSet.Tables("Users").Select(String.Format("RFID_UID = '{0}' AND IsActive = True", cleanUid.Replace("'", "''")))
            If rows.Length > 0 Then
                _terminalFailedTaps = 0
                Return rows(0)
            Else
                _terminalFailedTaps += 1
                If _terminalFailedTaps >= 5 Then
                    _terminalLockoutUntilUTC = DateTime.UtcNow.AddMinutes(5)
                End If
                Return Nothing
            End If
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

        Public Shared Function AddDocument(code As String, docType As String, title As String, origin As String, dest As String, cab As String, shelf As String, box As String, url As String, status As String, assigned As String, Optional flowDirection As String = "INCOMING", Optional assignedSection As String = "", Optional targetDeadline As String = "", Optional punchlist As String = "", Optional lastAction As String = "", Optional externalControlNumber As String = "", Optional isOffline As Boolean = True) As Integer
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
                If dt.Columns.Contains("PendingSync") Then r("PendingSync") = isOffline
                dt.Rows.Add(r)
                MarkDirty()

                Dim docId = CInt(r("DocumentID"))
                AddRoutingLog(docId, origin, dest, assigned, "REGISTERED", "Document registered in OSG System.", isOffline)
                AddMovementLog(docId, "INCOMING", String.Format("{0}/{1}/{2}", cab, shelf, box), assigned, "Initial storage assignment.", isOffline)
                Return docId
            End SyncLock
        End Function

        Public Shared Sub AddDirective(docId As Integer, directive As String, assignedTo As String, notes As String, staffName As String, Optional isOffline As Boolean = True)
            SyncLock _syncLock
                Dim dt = DataSet.Tables("Directives")
                Dim r = dt.NewRow()
                r("DocumentID") = docId
                r("SGDirective") = directive
                r("AssignedTo") = assignedTo
                r("Notes") = notes
                r("LogUser") = staffName
                r("Timestamp") = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                If dt.Columns.Contains("PendingSync") Then r("PendingSync") = isOffline
                dt.Rows.Add(r)

                Dim docRows = DataSet.Tables("Documents").Select("DocumentID = " & docId)
                If docRows.Length > 0 Then
                    If directive = "REVISION_REQUESTED" Then
                        docRows(0)("CurrentStatus") = "FOR_REVISION"
                    ElseIf directive = "APPROVED" Then
                        docRows(0)("CurrentStatus") = "APPROVED"
                    Else
                        docRows(0)("CurrentStatus") = "SG Directive: " & directive
                    End If
                    If Not String.IsNullOrEmpty(assignedTo) Then docRows(0)("AssignedStaff") = assignedTo
                    If isOffline AndAlso docRows(0).Table.Columns.Contains("PendingSync") Then
                        docRows(0)("PendingSync") = True
                    End If
                End If
                MarkDirty()
            End SyncLock
        End Sub

        Public Shared Sub AddRoutingLog(docId As Integer, fromOffice As String, toOffice As String, routedBy As String, action As String, remarks As String, Optional isOffline As Boolean = True)
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
                If dt.Columns.Contains("PendingSync") Then r("PendingSync") = isOffline
                dt.Rows.Add(r)

                Dim docRows = DataSet.Tables("Documents").Select("DocumentID = " & docId)
                If docRows.Length > 0 Then
                    docRows(0)("OriginatingOffice") = fromOffice
                    docRows(0)("DestinationOffice") = toOffice
                    If Not action.Equals("REVISION_REQUESTED", StringComparison.OrdinalIgnoreCase) Then
                        docRows(0)("LastActionTaken") = "Routed to " & toOffice & ": " & action
                    End If
                    If isOffline AndAlso docRows(0).Table.Columns.Contains("PendingSync") Then
                        docRows(0)("PendingSync") = True
                    End If
                End If
                MarkDirty()
            End SyncLock
        End Sub

        Public Shared Sub AddMovementLog(docId As Integer, fromLoc As String, toLoc As String, movedBy As String, reason As String, Optional isOffline As Boolean = True)
            SyncLock _syncLock
                Dim dt = DataSet.Tables("Movements")
                Dim r = dt.NewRow()
                r("DocumentID") = docId
                r("FromLocation") = fromLoc
                r("ToLocation") = toLoc
                r("MovedBy") = movedBy
                r("Reason") = reason
                r("Timestamp") = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
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
                    If isOffline AndAlso docRows(0).Table.Columns.Contains("PendingSync") Then
                        docRows(0)("PendingSync") = True
                    End If
                End If
                MarkDirty()
            End SyncLock
        End Sub

        Public Shared Sub LogAudit(user As String, action As String, Optional isOffline As Boolean = True)
            SyncLock _syncLock
                Dim dt = DataSet.Tables("AuditTrail")
                Dim r = dt.NewRow()
                r("UserName") = user
                r("ActionDescription") = action
                r("Timestamp") = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                If dt.Columns.Contains("PendingSync") Then r("PendingSync") = isOffline
                dt.Rows.Add(r)
                MarkDirty()
            End SyncLock
        End Sub

        Public Shared Sub RequestRevision(docId As Integer, punchlistNotes As String, returnSection As String, staffName As String)
            Dim extCnRev As String = ""
            SyncLock _syncLock
                EnsureInitialized()
                Dim docRows = DataSet.Tables("Documents").Select("DocumentID = " & docId)
                If docRows.Length = 0 Then Return
                Dim doc = docRows(0)

                doc("CurrentStatus") = "FOR_REVISION"
                Dim prevPunchlist As String = doc("RevisionPunchlist").ToString()
                Dim stamp As String = "[" & DateTime.Now.ToString("yyyy-MM-dd HH:mm") & "] " & punchlistNotes
                doc("RevisionPunchlist") = If(String.IsNullOrWhiteSpace(prevPunchlist), stamp, prevPunchlist & vbCrLf & stamp)

                Dim targetSec As String = If(Not String.IsNullOrWhiteSpace(returnSection), returnSection, DocumentService.GetDefaultSectionForCategory(doc("DocType").ToString()))
                doc("AssignedSection") = targetSec
                doc("LastActionTaken") = "Sec Gen requested revision: " & punchlistNotes
                MarkDirty()

                AddDirective(docId, "REVISION_REQUESTED", targetSec, punchlistNotes, staffName)
                AddRoutingLog(docId, "Office of the Secretary-General", targetSec, staffName, "REVISION_REQUESTED", punchlistNotes)
                LogAudit(staffName, "Revision requested for Document #" & docId & " with punchlist: " & punchlistNotes)

                extCnRev = If(doc.Table.Columns.Contains("ExternalControlNumber") AndAlso Not IsDBNull(doc("ExternalControlNumber")), doc("ExternalControlNumber").ToString(), "")
            End SyncLock

            If Not String.IsNullOrWhiteSpace(extCnRev) Then
                DocumentService.PushPortalStatusSafe(extCnRev, "FOR_REVISION")
            End If
        End Sub

        Public Shared Sub ResubmitDocument(docId As Integer, staffName As String, notes As String)
            Dim extCnResub As String = ""
            SyncLock _syncLock
                EnsureInitialized()
                Dim docRows = DataSet.Tables("Documents").Select("DocumentID = " & docId)
                If docRows.Length = 0 Then Return
                Dim doc = docRows(0)

                Dim originSec As String = doc("AssignedSection").ToString()
                doc("CurrentStatus") = "FOR_REVIEW"
                doc("AssignedSection") = "Secretary-General"
                doc("LastActionTaken") = "Resubmitted for Sec Gen review: " & notes
                MarkDirty()

                AddRoutingLog(docId, originSec, "Office of the Secretary-General", staffName, "RESUBMITTED", notes)
                LogAudit(staffName, "Document #" & docId & " resubmitted by " & originSec & " to Sec Gen: " & notes)

                extCnResub = If(doc.Table.Columns.Contains("ExternalControlNumber") AndAlso Not IsDBNull(doc("ExternalControlNumber")), doc("ExternalControlNumber").ToString(), "")
            End SyncLock

            If Not String.IsNullOrWhiteSpace(extCnResub) Then
                DocumentService.PushPortalStatusSafe(extCnResub, "FOR_REVIEW")
            End If
        End Sub

        Public Shared Sub ApproveDocument(docId As Integer, staffName As String, notes As String)
            Dim extCnAppr As String = ""
            SyncLock _syncLock
                EnsureInitialized()
                Dim docRows = DataSet.Tables("Documents").Select("DocumentID = " & docId)
                If docRows.Length = 0 Then Return
                Dim doc = docRows(0)

                doc("CurrentStatus") = "APPROVED"
                doc("AssignedSection") = "Records Section"
                doc("LastActionTaken") = "Approved by Secretary-General: " & notes
                MarkDirty()

                AddDirective(docId, "APPROVED", "Records Section", notes, staffName)
                AddRoutingLog(docId, "Office of the Secretary-General", "Records Section", staffName, "APPROVED", notes)
                LogAudit(staffName, "Document #" & docId & " approved by Sec Gen: " & notes)

                extCnAppr = If(doc.Table.Columns.Contains("ExternalControlNumber") AndAlso Not IsDBNull(doc("ExternalControlNumber")), doc("ExternalControlNumber").ToString(), "")
            End SyncLock

            If Not String.IsNullOrWhiteSpace(extCnAppr) Then
                DocumentService.PushPortalStatusSafe(extCnAppr, "APPROVED")
            End If
        End Sub

        Public Shared Sub ReleaseDocument(docId As Integer, staffName As String, notes As String)
            Dim extCnRel As String = ""
            SyncLock _syncLock
                EnsureInitialized()
                Dim docRows = DataSet.Tables("Documents").Select("DocumentID = " & docId)
                If docRows.Length = 0 Then Return
                Dim doc = docRows(0)

                Dim dest As String = doc("DestinationOffice").ToString()
                doc("CurrentStatus") = "RELEASED"
                doc("AssignedSection") = "Archived / Released"
                doc("LastActionTaken") = "Released to " & dest & ": " & notes
                MarkDirty()

                AddRoutingLog(docId, "Records Section", dest, staffName, "RELEASED", notes)
                LogAudit(staffName, "Document #" & docId & " released: " & notes)

                extCnRel = If(doc.Table.Columns.Contains("ExternalControlNumber") AndAlso Not IsDBNull(doc("ExternalControlNumber")), doc("ExternalControlNumber").ToString(), "")
            End SyncLock

            If Not String.IsNullOrWhiteSpace(extCnRel) Then
                DocumentService.PushPortalStatusSafe(extCnRel, "RELEASED")
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
            Dim docId = If(row.Table.Columns.Contains("DocumentID") AndAlso Not IsDBNull(row("DocumentID")), Convert.ToInt32(row("DocumentID")), 0)
            Dim docObj As New Document With {
                .DocumentID = docId,
                .DocCode = If(row.Table.Columns.Contains("DocCode") AndAlso Not IsDBNull(row("DocCode")), row("DocCode").ToString(), ""),
                .Title = If(row.Table.Columns.Contains("Title") AndAlso Not IsDBNull(row("Title")), row("Title").ToString(), ""),
                .OriginOffice = If(row.Table.Columns.Contains("OriginatingOffice") AndAlso Not IsDBNull(row("OriginatingOffice")), row("OriginatingOffice").ToString(), ""),
                .DestinationOffice = If(row.Table.Columns.Contains("DestinationOffice") AndAlso Not IsDBNull(row("DestinationOffice")), row("DestinationOffice").ToString(), ""),
                .FlowDirection = If(row.Table.Columns.Contains("FlowDirection") AndAlso Not IsDBNull(row("FlowDirection")), row("FlowDirection").ToString(), "INCOMING"),
                .AssignedSection = If(row.Table.Columns.Contains("AssignedSection") AndAlso Not IsDBNull(row("AssignedSection")), row("AssignedSection").ToString(), ""),
                .RevisionPunchlist = If(row.Table.Columns.Contains("RevisionPunchlist") AndAlso Not IsDBNull(row("RevisionPunchlist")), row("RevisionPunchlist").ToString(), ""),
                .LastActionTaken = If(row.Table.Columns.Contains("LastActionTaken") AndAlso Not IsDBNull(row("LastActionTaken")), row("LastActionTaken").ToString(), ""),
                .RegisteredAtUTC = DateTime.UtcNow
            }

            If row.Table.Columns.Contains("Remarks") AndAlso Not IsDBNull(row("Remarks")) Then
                docObj.Remarks = row("Remarks").ToString()
            End If

            Dim typeStr = If(row.Table.Columns.Contains("DocType") AndAlso Not IsDBNull(row("DocType")), row("DocType").ToString().ToUpperInvariant(), "")
            Select Case typeStr
                Case "FINANCE", "FIN": docObj.DocumentTypeID = 3
                Case "TRAVEL", "TRAVEL ORDER", "TO": docObj.DocumentTypeID = 4
                Case "LEGISLATIVE", "LEG": docObj.DocumentTypeID = 2
                Case Else: docObj.DocumentTypeID = 1
            End Select

            Dim statusStr = If(row.Table.Columns.Contains("CurrentStatus") AndAlso Not IsDBNull(row("CurrentStatus")), row("CurrentStatus").ToString().ToUpperInvariant(), "")
            Select Case statusStr
                Case "RECEIVED", "LOGGED": docObj.StatusID = 1
                Case "FOR_REVIEW": docObj.StatusID = 2
                Case "FOR_REVISION": docObj.StatusID = 3
                Case "APPROVED": docObj.StatusID = 6
                Case "RELEASED": docObj.StatusID = 7
                Case "FILED", "ARCHIVED": docObj.StatusID = 8
                Case Else: docObj.StatusID = 1
            End Select

            If row.Table.Columns.Contains("TargetDeadlineUTC") AndAlso Not IsDBNull(row("TargetDeadlineUTC")) Then
                Dim dVal As DateTime
                If DateTime.TryParse(row("TargetDeadlineUTC").ToString(), dVal) Then
                    docObj.TargetDeadlineUTC = dVal
                End If
            End If

            If row.Table.Columns.Contains("CurrentStorageLocationID") AndAlso Not IsDBNull(row("CurrentStorageLocationID")) Then
                Dim sId As Integer
                If Integer.TryParse(row("CurrentStorageLocationID").ToString(), sId) Then
                    docObj.CurrentStorageLocationID = sId
                End If
            End If

            If row.Table.Columns.Contains("CabinetID") AndAlso Not IsDBNull(row("CabinetID")) Then docObj.CabinetID = row("CabinetID").ToString()
            If row.Table.Columns.Contains("ShelfNo") AndAlso Not IsDBNull(row("ShelfNo")) Then docObj.ShelfNo = row("ShelfNo").ToString()
            If row.Table.Columns.Contains("BoxCode") AndAlso Not IsDBNull(row("BoxCode")) Then docObj.BoxCode = row("BoxCode").ToString()
            If row.Table.Columns.Contains("ExternalControlNumber") AndAlso Not IsDBNull(row("ExternalControlNumber")) Then docObj.ExternalControlNumber = row("ExternalControlNumber").ToString()

            Return docObj
        End Function

        Public Shared Function GetDocumentByID(docId As Integer) As Document
            SyncLock _syncLock
                EnsureInitialized()
                Dim dt = DataSet.Tables("Documents")
                If dt Is Nothing Then Return Nothing
                Dim rows = dt.Select("DocumentID = " & docId)
                If rows.Length = 0 Then Return Nothing
                Return MapRowToDocument(rows(0))
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
                    If r.Table.Columns.Contains("Timestamp") AndAlso Not IsDBNull(r("Timestamp")) Then
                        Dim dtVal As DateTime
                        If DateTime.TryParse(r("Timestamp").ToString(), dtVal) Then
                            logEntry.RoutedAtUTC = dtVal
                        End If
                    End If
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
                    If r.Table.Columns.Contains("Timestamp") AndAlso Not IsDBNull(r("Timestamp")) Then
                        Dim dtVal As DateTime
                        If DateTime.TryParse(r("Timestamp").ToString(), dtVal) Then
                            dirEntry.IssuedAtUTC = dtVal
                        End If
                    End If
                    list.Add(dirEntry)
                Next
                Return list
            End SyncLock
        End Function
    End Class
End Namespace
