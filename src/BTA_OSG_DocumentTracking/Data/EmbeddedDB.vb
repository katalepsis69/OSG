Imports System
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
                    EnsureDefaultUsers()
                    Return
                Catch ex As Exception
                    ' Fallback to re-creating if corrupt
                End Try
            End If

            SeedDefaultData()
            Save()
        End Sub

        Public Shared Sub Save()
            Try
                DataSet.WriteXml(DbPath, XmlWriteMode.WriteSchema)
            Catch ex As Exception
            End Try
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
            dtUsers.Columns.Add("IsActive", GetType(Boolean))
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
            dtAudit.PrimaryKey = New DataColumn() {dtAudit.Columns("AuditID")}
            DataSet.Tables.Add(dtAudit)
        End Sub

        Private Shared Sub EnsureDefaultUsers()
            Dim dtUsers = DataSet.Tables("Users")
            If dtUsers.Rows.Count = 0 Then
                AddUser("88A9F321", "Prof. Ali B. Pangalian", "Secretary-General")
                AddUser("99B1C456", "Atty. Fatima Z. Rasheed", "OSG Chief")
                AddUser("77C3D987", "Omire Khalid B. Ebrahim", "System Administrator")
                AddUser("55E5F666", "Hassim A. Ibrahim", "Administrative Staff")
                AddUser("11A2B3C4", "CJ Fairoz A. Usop", "Administrative Staff")
            End If
        End Sub

        Private Shared Sub SeedDefaultData()
            EnsureDefaultUsers()

            Dim id1 = AddDocument("RES-2026-001", "Resolution", "Resolution Expressing Gratitude to BARMM Chief Minister", "Office of MP Yasser", "Office of Secretary-General", "CAB-A", "S-1", "BOX-01", "https://drive.google.com/file/d/sample-res-001/view", "LOGGED", "Hassim A. Ibrahim")
            AddDirective(id1, "For Immediate Action", "Hassim A. Ibrahim", "Priority processing requested", "Prof. Ali B. Pangalian")

            Dim id2 = AddDocument("BLL-2026-001", "Parliament Bill", "Bangsamoro Education Code Amendment Bill of 2026", "Committee on Education", "Committee on Rules", "CAB-B", "S-3", "BOX-04", "https://drive.google.com/file/d/sample-bll-001/view", "LOGGED", "CJ Fairoz A. Usop")
            AddDirective(id2, "Referred to Committee on Rules", "CJ Fairoz A. Usop", "Referral per SG instruction", "Atty. Fatima Z. Rasheed")

            Dim id3 = AddDocument("REP-2026-001", "Committee Report", "Report on BTA Budgetary Allocations for BARMM Infrastructure", "Committee on Finance", "Speaker's Office", "CAB-C", "S-2", "BOX-02", "https://drive.google.com/file/d/sample-rep-001/view", "LOGGED", "Atty. Fatima Z. Rasheed")
            AddDirective(id3, "Forwarded for Speaker Signature", "Atty. Fatima Z. Rasheed", "Awaiting signature", "Prof. Ali B. Pangalian")

            LogAudit("SYSTEM", "Embedded Database Engine initialized with Parliamentary seed dataset.")
        End Sub

        Public Shared Sub AddUser(uid As String, name As String, role As String)
            Dim dt = DataSet.Tables("Users")
            Dim cleanUid = uid.Trim().ToUpperInvariant()
            Dim existing = dt.Select(String.Format("RFID_UID = '{0}'", cleanUid.Replace("'", "''")))
            If existing.Length > 0 Then
                existing(0)("FullName") = name
                existing(0)("Role") = role
                existing(0)("IsActive") = True
            Else
                Dim r = dt.NewRow()
                r("RFID_UID") = cleanUid
                r("FullName") = name
                r("Role") = role
                r("IsActive") = True
                dt.Rows.Add(r)
            End If
            Save()
        End Sub

        Public Shared Function AuthenticateRFID(uid As String) As DataRow
            If String.IsNullOrWhiteSpace(uid) Then Return Nothing
            Dim cleanUid = uid.Trim().ToUpperInvariant()
            Dim rows = DataSet.Tables("Users").Select(String.Format("RFID_UID = '{0}' AND IsActive = True", cleanUid.Replace("'", "''")))
            If rows.Length > 0 Then Return rows(0)
            Return Nothing
        End Function

        Public Shared Function GenerateDocCode(docType As String) As String
            Dim prefix As String = "DOC"
            Select Case docType.Trim().ToLower()
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
            If Not trimmed.StartsWith("https://", StringComparison.OrdinalIgnoreCase) Then
                errMessage = "URL must use secure 'https://' protocol."
                Return False
            End If
            Try
                Dim uri As New Uri(trimmed)
                Dim host = uri.Host.ToLower()
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

        Public Shared Function AddDocument(code As String, docType As String, title As String, origin As String, dest As String, cab As String, shelf As String, box As String, url As String, status As String, assigned As String) As Integer
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
            dt.Rows.Add(r)
            Save()

            Dim docId = CInt(r("DocumentID"))
            AddRoutingLog(docId, origin, dest, assigned, "REGISTERED", "Document registered in OSG System.")
            AddMovementLog(docId, "INCOMING", String.Format("{0}/{1}/{2}", cab, shelf, box), assigned, "Initial storage assignment.")
            Return docId
        End Function

        Public Shared Sub AddDirective(docId As Integer, directive As String, assignedTo As String, notes As String, staffName As String)
            Dim dt = DataSet.Tables("Directives")
            Dim r = dt.NewRow()
            r("DocumentID") = docId
            r("SGDirective") = directive
            r("AssignedTo") = assignedTo
            r("Notes") = notes
            r("LogUser") = staffName
            r("Timestamp") = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            dt.Rows.Add(r)

            Dim docRows = DataSet.Tables("Documents").Select("DocumentID = " & docId)
            If docRows.Length > 0 Then
                docRows(0)("CurrentStatus") = "SG Directive: " & directive
                If Not String.IsNullOrEmpty(assignedTo) Then docRows(0)("AssignedStaff") = assignedTo
            End If
            Save()
        End Sub

        Public Shared Sub AddRoutingLog(docId As Integer, fromOffice As String, toOffice As String, routedBy As String, action As String, remarks As String)
            Dim dt = DataSet.Tables("RoutingLogs")
            Dim r = dt.NewRow()
            r("DocumentID") = docId
            r("FromOffice") = fromOffice
            r("ToOffice") = toOffice
            r("RoutedBy") = routedBy
            r("ActionTaken") = action
            r("Remarks") = remarks
            r("Timestamp") = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            dt.Rows.Add(r)

            Dim docRows = DataSet.Tables("Documents").Select("DocumentID = " & docId)
            If docRows.Length > 0 Then
                docRows(0)("OriginatingOffice") = fromOffice
                docRows(0)("DestinationOffice") = toOffice
            End If
            Save()
        End Sub

        Public Shared Sub AddMovementLog(docId As Integer, fromLoc As String, toLoc As String, movedBy As String, reason As String)
            Dim dt = DataSet.Tables("Movements")
            Dim r = dt.NewRow()
            r("DocumentID") = docId
            r("FromLocation") = fromLoc
            r("ToLocation") = toLoc
            r("MovedBy") = movedBy
            r("Reason") = reason
            r("Timestamp") = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            dt.Rows.Add(r)

            Dim parts = toLoc.Split("/"c)
            If parts.Length >= 3 Then
                Dim docRows = DataSet.Tables("Documents").Select("DocumentID = " & docId)
                If docRows.Length > 0 Then
                    docRows(0)("CabinetID") = parts(0).Trim()
                    docRows(0)("ShelfNo") = parts(1).Trim()
                    docRows(0)("BoxCode") = parts(2).Trim()
                End If
            End If
            Save()
        End Sub

        Public Shared Sub LogAudit(user As String, action As String)
            Dim dt = DataSet.Tables("AuditTrail")
            Dim r = dt.NewRow()
            r("UserName") = user
            r("ActionDescription") = action
            r("Timestamp") = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            dt.Rows.Add(r)
            Save()
        End Sub

        Public Shared Function GetVisibleDocuments(user As DataRow) As DataTable
            Dim dt = DataSet.Tables("Documents")
            If user Is Nothing Then Return dt.Clone()

            Dim role As String = user("Role").ToString()
            If role = "Secretary-General" OrElse role = "System Administrator" OrElse role = "OSG Chief" Then
                Return dt
            End If

            Dim staffName As String = user("FullName").ToString().Replace("'", "''")
            Dim dv As New DataView(dt)
            dv.RowFilter = String.Format("AssignedStaff = '{0}' OR AssignedStaff = '' OR AssignedStaff IS NULL", staffName)
            Return dv.ToTable()
        End Function
    End Class
End Namespace
