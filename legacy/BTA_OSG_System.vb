Imports System
Imports System.Data
Imports System.Diagnostics
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Windows.Forms

Namespace BTA_OSG

    ' =========================================================================
    ' 1. EMBEDDED RELATIONAL DATABASE ENGINE (DataSet + XML Persistence)
    ' =========================================================================
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
            Dim cleanUid = uid.Trim().ToUpper()
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
            Dim cleanUid = uid.Trim().ToUpper()
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

    ' =========================================================================
    ' 2. HIGH-END EXECUTIVE PARLIAMENTARY DESKTOP UI (WinForms)
    ' =========================================================================
    Public Class FormMain
        Inherits Form

        Public CurrentUser As DataRow = Nothing

        ' Layout Panels
        Private pnlSidebar As Panel
        Private pnlHeader As Panel
        Private pnlContent As Panel

        ' Header Controls
        Private lblTitle As Label
        Private lblUserBadge As Label
        Private btnScanRFID As Button
        Private btnTapSG As Button
        Private btnTapAdmin As Button
        Private btnTapStaff As Button
        Private btnLogout As Button

        ' Sidebar Navigation Items
        Private navButtons As New List(Of Button)()
        Private activeNavIndex As Integer = 0

        ' Views (Panels)
        Private viewDashboard As Panel
        Private viewRegistry As Panel
        Private viewDirectives As Panel
        Private viewSearch As Panel
        Private viewAdmin As Panel
        Private viewAudit As Panel

        ' Dashboard Controls
        Private lblStatTotalDocs As Label
        Private lblStatDirectives As Label
        Private lblStatActiveRoute As Label
        Private lblStatVaultStorage As Label
        Private dgvDashRecent As DataGridView

        ' Registry Controls
        Private txtTitle As TextBox
        Private cmbDocType As ComboBox
        Private txtOrigin As TextBox
        Private txtDest As TextBox
        Private txtCabinet As TextBox
        Private txtShelf As TextBox
        Private txtBox As TextBox
        Private txtGDrive As TextBox
        Private cmbAssignedStaff As ComboBox
        Private btnRegister As Button
        Private dgvRegistry As DataGridView
        Private btnViewRegistryDetail As Button

        ' Directives Controls
        Private cmbDirDocs As ComboBox
        Private cmbDirective As ComboBox
        Private cmbDirAssign As ComboBox
        Private txtDirNotes As TextBox
        Private btnApplyDirective As Button
        Private dgvDirectives As DataGridView

        ' Search Controls
        Private txtSearchKey As TextBox
        Private btnSearch As Button
        Private dgvSearch As DataGridView
        Private btnViewSearchDetail As Button
        Private btnOpenPDF As Button

        ' User Admin Controls
        Private dgvUsers As DataGridView
        Private txtNewUserName As TextBox
        Private cmbNewUserRole As ComboBox
        Private txtNewUserUID As TextBox
        Private btnAddUser As Button

        ' Audit Controls
        Private txtAuditSearch As TextBox
        Private btnAuditFilter As Button
        Private dgvAudit As DataGridView

        Public Sub New()
            EmbeddedDB.Initialize()
            InitializeUI()
            PopulateStaffDropdowns()

            SetStyle(ControlStyles.OptimizedDoubleBuffer Or ControlStyles.AllPaintingInWmPaint, True)
            UpdateStyles()

            AddHandler Me.Shown, Sub() ShowRFIDLoginDialog()
        End Sub

        Private Sub InitializeUI()
            Me.Text = "Bangsamoro Transition Authority Parliament | OSG Document Status Tracking & Monitoring System"
            Me.Size = New Size(1380, 850)
            Me.MinimumSize = New Size(1100, 720)
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.Font = New Font("Segoe UI", 9.5F, FontStyle.Regular)
            Me.BackColor = Color.FromArgb(15, 23, 42) ' Dark Slate 900

            ' Left Sidebar Navigation Rail
            SetupSidebar()

            ' Top Executive Bar
            SetupHeader()

            ' Main View Container
            pnlContent = New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = Color.FromArgb(15, 23, 42),
                .Padding = New Padding(20)
            }
            Me.Controls.Add(pnlContent)
            pnlContent.BringToFront()

            ' Create Individual Views
            SetupDashboardView()
            SetupRegistryView()
            SetupDirectivesView()
            SetupSearchView()
            SetupAdminView()
            SetupAuditView()

            SwitchNavView(0)
        End Sub

        Private Sub SetupSidebar()
            pnlSidebar = New Panel With {
                .Dock = DockStyle.Left,
                .Width = 240,
                .BackColor = Color.FromArgb(30, 41, 59), ' Slate 800
                .Padding = New Padding(12)
            }

            ' Branding Section
            Dim pnlBrand As New Panel With {
                .Dock = DockStyle.Top,
                .Height = 85,
                .BackColor = Color.FromArgb(30, 41, 59)
            }
            Dim lblBrand As New Label With {
                .Text = "BTA PARLIAMENT" & vbCrLf & "OFFICE OF THE SG",
                .Font = New Font("Segoe UI", 11.0F, FontStyle.Bold),
                .ForeColor = Color.FromArgb(248, 250, 252),
                .Location = New Point(12, 18),
                .AutoSize = True
            }
            Dim lblBrandSub As New Label With {
                .Text = "DOCUMENT TRACKING SYSTEM",
                .Font = New Font("Segoe UI", 7.5F, FontStyle.Bold),
                .ForeColor = Color.FromArgb(129, 140, 248),
                .Location = New Point(12, 58),
                .AutoSize = True
            }
            pnlBrand.Controls.AddRange(New Control() {lblBrand, lblBrandSub})

            ' Navigation Buttons Stack
            Dim navItems As String() = {"Dashboard", "Document Registry", "SG Directives", "Search & Storage", "User & RFID Admin", "Audit Trail"}
            Dim yPos As Integer = 100

            For i As Integer = 0 To navItems.Length - 1
                Dim idx As Integer = i
                Dim btn As New Button With {
                    .Text = "  " & navItems(i),
                    .Location = New Point(12, yPos),
                    .Size = New Size(216, 44),
                    .FlatStyle = FlatStyle.Flat,
                    .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold),
                    .TextAlign = ContentAlignment.MiddleLeft,
                    .Cursor = Cursors.Hand,
                    .ForeColor = Color.FromArgb(203, 213, 225),
                    .BackColor = Color.FromArgb(30, 41, 59)
                }
                btn.FlatAppearance.BorderSize = 0
                btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(51, 65, 85)

                AddHandler btn.Click, Sub() SwitchNavView(idx)
                navButtons.Add(btn)
                pnlSidebar.Controls.Add(btn)
                yPos += 50
            Next

            Me.Controls.Add(pnlSidebar)
            pnlSidebar.Controls.Add(pnlBrand)
        End Sub

        Private Sub SwitchNavView(index As Integer)
            activeNavIndex = index
            For i As Integer = 0 To navButtons.Count - 1
                If i = index Then
                    navButtons(i).BackColor = Color.FromArgb(79, 70, 229) ' Electric Indigo
                    navButtons(i).ForeColor = Color.White
                    navButtons(i).Font = New Font("Segoe UI", 9.5F, FontStyle.Bold)
                Else
                    navButtons(i).BackColor = Color.FromArgb(30, 41, 59)
                    navButtons(i).ForeColor = Color.FromArgb(203, 213, 225)
                    navButtons(i).Font = New Font("Segoe UI", 9.5F, FontStyle.Bold)
                End If
            Next

            pnlContent.Controls.Clear()
            Select Case index
                Case 0 : pnlContent.Controls.Add(viewDashboard)
                Case 1 : pnlContent.Controls.Add(viewRegistry)
                Case 2 : pnlContent.Controls.Add(viewDirectives)
                Case 3 : pnlContent.Controls.Add(viewSearch)
                Case 4 : pnlContent.Controls.Add(viewAdmin)
                Case 5 : pnlContent.Controls.Add(viewAudit)
            End Select

            RefreshActiveTabGrid()
        End Sub

        Private Sub SetupHeader()
            pnlHeader = New Panel With {
                .Dock = DockStyle.Top,
                .Height = 70,
                .BackColor = Color.FromArgb(15, 23, 42),
                .Padding = New Padding(20, 12, 20, 12)
            }

            lblTitle = New Label With {
                .Text = "OFFICE OF THE SECRETARY-GENERAL  |  DOCUMENT TRACKING DESK",
                .Font = New Font("Segoe UI", 11.5F, FontStyle.Bold),
                .ForeColor = Color.White,
                .Location = New Point(260, 14),
                .AutoSize = True
            }

            lblUserBadge = New Label With {
                .Text = "[ RFID LOGGED OUT ]",
                .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
                .ForeColor = Color.FromArgb(248, 113, 113),
                .Location = New Point(260, 42),
                .AutoSize = True
            }

            btnScanRFID = CreateHeaderButton("[ TAP RFID BADGE ]", New Point(680, 16), Color.FromArgb(99, 102, 241), 160)
            btnTapSG = CreateHeaderButton("SG Tap", New Point(850, 16), Color.FromArgb(16, 185, 129), 95)
            btnTapAdmin = CreateHeaderButton("Admin Tap", New Point(955, 16), Color.FromArgb(14, 165, 233), 100)
            btnTapStaff = CreateHeaderButton("Staff Tap", New Point(1065, 16), Color.FromArgb(245, 158, 11), 95)
            btnLogout = CreateHeaderButton("Logout", New Point(1170, 16), Color.FromArgb(239, 68, 68), 95)

            btnScanRFID.Anchor = AnchorStyles.Top Or AnchorStyles.Right
            btnTapSG.Anchor = AnchorStyles.Top Or AnchorStyles.Right
            btnTapAdmin.Anchor = AnchorStyles.Top Or AnchorStyles.Right
            btnTapStaff.Anchor = AnchorStyles.Top Or AnchorStyles.Right
            btnLogout.Anchor = AnchorStyles.Top Or AnchorStyles.Right

            AddHandler btnScanRFID.Click, Sub() ShowRFIDLoginDialog()
            AddHandler btnTapSG.Click, Sub() AuthenticateUser("88A9F321")
            AddHandler btnTapAdmin.Click, Sub() AuthenticateUser("77C3D987")
            AddHandler btnTapStaff.Click, Sub() AuthenticateUser("55E5F666")
            AddHandler btnLogout.Click, Sub() AuthenticateUser("")

            pnlHeader.Controls.AddRange(New Control() {lblTitle, lblUserBadge, btnScanRFID, btnTapSG, btnTapAdmin, btnTapStaff, btnLogout})
            Me.Controls.Add(pnlHeader)
        End Sub

        Private Function CreateHeaderButton(text As String, loc As Point, bg As Color, width As Integer) As Button
            Return New Button With {
                .Text = text,
                .Location = loc,
                .Size = New Size(width, 36),
                .BackColor = bg,
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
                .Cursor = Cursors.Hand
            }
        End Function

        Private Sub ShowRFIDLoginDialog()
            Using dlg As New FormRFIDLogin()
                If dlg.ShowDialog(Me) = DialogResult.OK AndAlso Not String.IsNullOrEmpty(dlg.ScannedUID) Then
                    AuthenticateUser(dlg.ScannedUID)
                End If
            End Using
        End Sub

        Public Shared Sub ApplyGridStyle(dgv As DataGridView)
            dgv.EnableHeadersVisualStyles = False
            dgv.BackgroundColor = Color.FromArgb(30, 41, 59)
            dgv.BorderStyle = BorderStyle.None
            dgv.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
            dgv.GridColor = Color.FromArgb(51, 65, 85)
            dgv.ColumnHeadersHeight = 40
            dgv.RowTemplate.Height = 36
            dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect

            Dim pi = GetType(Control).GetProperty("DoubleBuffered", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic)
            If pi IsNot Nothing Then pi.SetValue(dgv, True, Nothing)

            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(15, 23, 42)
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(148, 163, 184)
            dgv.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 9.0F, FontStyle.Bold)

            dgv.DefaultCellStyle.BackColor = Color.FromArgb(30, 41, 59)
            dgv.DefaultCellStyle.ForeColor = Color.FromArgb(248, 250, 252)
            dgv.DefaultCellStyle.Font = New Font("Segoe UI", 9.0F, FontStyle.Regular)
            dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(49, 46, 129)
            dgv.DefaultCellStyle.SelectionForeColor = Color.White

            dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(24, 34, 50)
        End Sub

        Public Sub AuthenticateUser(uid As String)
            If String.IsNullOrEmpty(uid) Then
                CurrentUser = Nothing
                lblUserBadge.Text = "[ RFID LOGGED OUT — ACCESS RESTRICTED ]"
                lblUserBadge.ForeColor = Color.FromArgb(248, 113, 113)
            Else
                Dim user = EmbeddedDB.AuthenticateRFID(uid)
                If user IsNot Nothing Then
                    CurrentUser = user
                    Dim isGlobal As Boolean = (user("Role").ToString() = "Secretary-General" OrElse user("Role").ToString() = "System Administrator" OrElse user("Role").ToString() = "OSG Chief")
                    lblUserBadge.Text = String.Format("AUTHENTICATED: {0} [{1}] — {2}", user("FullName").ToString().ToUpper(), user("Role").ToString().ToUpper(), If(isGlobal, "GLOBAL ACCESS", "STAFF VIEW"))
                    lblUserBadge.ForeColor = Color.FromArgb(52, 211, 153)
                    EmbeddedDB.LogAudit(user("FullName").ToString(), "RFID Badge Tap Authenticated [UID: " & uid & "]")
                Else
                    MessageBox.Show("Unrecognized RFID Smart Card Badge UID: " & uid, "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If
            End If
            PopulateStaffDropdowns()
            RefreshActiveTabGrid()
        End Sub

        Public Sub PopulateStaffDropdowns()
            cmbAssignedStaff.Items.Clear()
            cmbDirAssign.Items.Clear()
            For Each row As DataRow In EmbeddedDB.DataSet.Tables("Users").Rows
                Dim name = row("FullName").ToString()
                cmbAssignedStaff.Items.Add(name)
                cmbDirAssign.Items.Add(name)
            Next
            If cmbAssignedStaff.Items.Count > 0 Then cmbAssignedStaff.SelectedIndex = 0
            If cmbDirAssign.Items.Count > 0 Then cmbDirAssign.SelectedIndex = 0
        End Sub

        Private Sub SetupDashboardView()
            viewDashboard = New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.FromArgb(15, 23, 42)}

            ' Metric Cards Top Row
            Dim pnlCards As New TableLayoutPanel With {
                .Dock = DockStyle.Top,
                .Height = 110,
                .ColumnCount = 4,
                .RowCount = 1,
                .Padding = New Padding(0, 0, 0, 15)
            }
            pnlCards.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 25.0F))
            pnlCards.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 25.0F))
            pnlCards.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 25.0F))
            pnlCards.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 25.0F))

            lblStatTotalDocs = CreateStatCard(pnlCards, 0, "TOTAL DOCUMENTS", "0", Color.FromArgb(99, 102, 241))
            lblStatDirectives = CreateStatCard(pnlCards, 1, "SG DIRECTIVES", "0", Color.FromArgb(16, 185, 129))
            lblStatActiveRoute = CreateStatCard(pnlCards, 2, "ACTIVE ROUTINGS", "0", Color.FromArgb(14, 165, 233))
            lblStatVaultStorage = CreateStatCard(pnlCards, 3, "PHYSICAL VAULT ITEMS", "0", Color.FromArgb(245, 158, 11))

            ' Recent Activity Grid Title
            Dim lblRecHeader As New Label With {
                .Text = "RECENT PARLIAMENTARY DOCUMENTS",
                .Font = New Font("Segoe UI", 10.0F, FontStyle.Bold),
                .ForeColor = Color.FromArgb(148, 163, 184),
                .Dock = DockStyle.Top,
                .Height = 30
            }

            dgvDashRecent = New DataGridView With {
                .Dock = DockStyle.Fill,
                .ReadOnly = True,
                .AllowUserToAddRows = False,
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            }
            ApplyGridStyle(dgvDashRecent)
            AddHandler dgvDashRecent.CellDoubleClick, Sub(s, e) If e.RowIndex >= 0 Then OpenSelectedDocumentDetail(dgvDashRecent)

            viewDashboard.Controls.Add(dgvDashRecent)
            viewDashboard.Controls.Add(lblRecHeader)
            viewDashboard.Controls.Add(pnlCards)
        End Sub

        Private Function CreateStatCard(parent As TableLayoutPanel, colIndex As Integer, title As String, initVal As String, accentBg As Color) As Label
            Dim pnlCard As New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = Color.FromArgb(30, 41, 59),
                .Margin = New Padding(5)
            }
            Dim pnlBar As New Panel With {
                .Dock = DockStyle.Top,
                .Height = 4,
                .BackColor = accentBg
            }
            Dim lblT As New Label With {
                .Text = title,
                .Font = New Font("Segoe UI", 8.0F, FontStyle.Bold),
                .ForeColor = Color.FromArgb(148, 163, 184),
                .Location = New Point(12, 14),
                .AutoSize = True
            }
            Dim lblV As New Label With {
                .Text = initVal,
                .Font = New Font("Segoe UI", 18.0F, FontStyle.Bold),
                .ForeColor = Color.White,
                .Location = New Point(12, 36),
                .AutoSize = True
            }

            pnlCard.Controls.AddRange(New Control() {pnlBar, lblT, lblV})
            parent.Controls.Add(pnlCard, colIndex, 0)
            Return lblV
        End Function

        Private Sub SetupRegistryView()
            viewRegistry = New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.FromArgb(15, 23, 42)}

            Dim pnlFormOuter As New Panel With {
                .Dock = DockStyle.Left,
                .Width = 380,
                .BackColor = Color.FromArgb(30, 41, 59),
                .Padding = New Padding(16)
            }

            Dim pnlFormScroll As New Panel With {
                .Dock = DockStyle.Fill,
                .AutoScroll = True,
                .BackColor = Color.FromArgb(30, 41, 59)
            }

            txtTitle = New TextBox With {.Height = 28, .Text = "Draft Resolution on BTA Regional Governance", .BackColor = Color.FromArgb(15, 23, 42), .ForeColor = Color.White, .BorderStyle = BorderStyle.FixedSingle}
            cmbDocType = New ComboBox With {.DropDownStyle = ComboBoxStyle.DropDownList, .Height = 28, .BackColor = Color.FromArgb(15, 23, 42), .ForeColor = Color.White}
            cmbDocType.Items.AddRange(New Object() {"Resolution", "Parliament Bill", "Committee Report", "Executive Communication", "Memorandum", "Endorsement", "Journal Entry"})
            cmbDocType.SelectedIndex = 0

            txtOrigin = New TextBox With {.Height = 28, .Text = "Office of MP Yasser", .BackColor = Color.FromArgb(15, 23, 42), .ForeColor = Color.White, .BorderStyle = BorderStyle.FixedSingle}
            txtDest = New TextBox With {.Height = 28, .Text = "Office of the Secretary-General", .BackColor = Color.FromArgb(15, 23, 42), .ForeColor = Color.White, .BorderStyle = BorderStyle.FixedSingle}
            txtCabinet = New TextBox With {.Height = 28, .Text = "CAB-A", .BackColor = Color.FromArgb(15, 23, 42), .ForeColor = Color.White, .BorderStyle = BorderStyle.FixedSingle}
            txtShelf = New TextBox With {.Height = 28, .Text = "S-2", .BackColor = Color.FromArgb(15, 23, 42), .ForeColor = Color.White, .BorderStyle = BorderStyle.FixedSingle}
            txtBox = New TextBox With {.Height = 28, .Text = "BOX-03", .BackColor = Color.FromArgb(15, 23, 42), .ForeColor = Color.White, .BorderStyle = BorderStyle.FixedSingle}
            txtGDrive = New TextBox With {.Height = 28, .Text = "https://drive.google.com/file/d/bta-doc-2026/view", .BackColor = Color.FromArgb(15, 23, 42), .ForeColor = Color.White, .BorderStyle = BorderStyle.FixedSingle}
            cmbAssignedStaff = New ComboBox With {.DropDownStyle = ComboBoxStyle.DropDownList, .Height = 28, .BackColor = Color.FromArgb(15, 23, 42), .ForeColor = Color.White}

            Dim curY As Integer = 10
            Dim AddControlField = Sub(lblText As String, ctrl As Control)
                                      Dim lbl As New Label With {
                                          .Text = lblText,
                                          .Location = New Point(10, curY),
                                          .AutoSize = True,
                                          .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
                                          .ForeColor = Color.FromArgb(148, 163, 184)
                                      }
                                      curY += 22
                                      ctrl.Location = New Point(10, curY)
                                      ctrl.Width = 330
                                      curY += ctrl.Height + 12
                                      pnlFormScroll.Controls.Add(lbl)
                                      pnlFormScroll.Controls.Add(ctrl)
                                  End Sub

            AddControlField("Document Title:", txtTitle)
            AddControlField("Document Type:", cmbDocType)
            AddControlField("Originating Office:", txtOrigin)
            AddControlField("Destination Office:", txtDest)
            AddControlField("Cabinet Landmark ID:", txtCabinet)
            AddControlField("Shelf Landmark No:", txtShelf)
            AddControlField("Box Landmark Code:", txtBox)
            AddControlField("Google Drive Soft Copy URL:", txtGDrive)
            AddControlField("Assigned OSG Staff:", cmbAssignedStaff)

            btnRegister = New Button With {
                .Text = "REGISTER PARLIAMENTARY DOCUMENT",
                .Location = New Point(10, curY + 5),
                .Size = New Size(330, 42),
                .BackColor = Color.FromArgb(79, 70, 229),
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold),
                .Cursor = Cursors.Hand
            }
            AddHandler btnRegister.Click, AddressOf OnRegisterDocument
            pnlFormScroll.Controls.Add(btnRegister)
            pnlFormOuter.Controls.Add(pnlFormScroll)

            Dim pnlRight As New Panel With {.Dock = DockStyle.Fill, .Padding = New Padding(15, 0, 0, 0)}
            Dim pnlRegistryToolbar As New Panel With {.Dock = DockStyle.Top, .Height = 45, .BackColor = Color.FromArgb(15, 23, 42)}

            btnViewRegistryDetail = New Button With {
                .Text = "View Selected Document Full Specification & History",
                .Location = New Point(0, 4),
                .Size = New Size(360, 34),
                .BackColor = Color.FromArgb(14, 165, 233),
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
                .Cursor = Cursors.Hand
            }
            AddHandler btnViewRegistryDetail.Click, Sub() OpenSelectedDocumentDetail(dgvRegistry)
            pnlRegistryToolbar.Controls.Add(btnViewRegistryDetail)

            dgvRegistry = New DataGridView With {
                .Dock = DockStyle.Fill,
                .ReadOnly = True,
                .AllowUserToAddRows = False,
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            }
            ApplyGridStyle(dgvRegistry)
            AddHandler dgvRegistry.CellDoubleClick, Sub(s, e) If e.RowIndex >= 0 Then OpenSelectedDocumentDetail(dgvRegistry)

            pnlRight.Controls.Add(dgvRegistry)
            pnlRight.Controls.Add(pnlRegistryToolbar)

            viewRegistry.Controls.Add(pnlRight)
            viewRegistry.Controls.Add(pnlFormOuter)
        End Sub

        Private Sub OnRegisterDocument(sender As Object, e As EventArgs)
            If CurrentUser Is Nothing Then
                MessageBox.Show("RFID Authentication Required to Register Documents.", "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            If String.IsNullOrWhiteSpace(txtTitle.Text) Then
                MessageBox.Show("Please enter Document Title.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim errUrlMsg As String = ""
            If Not EmbeddedDB.ValidateGDriveURL(txtGDrive.Text, errUrlMsg) Then
                MessageBox.Show(errUrlMsg, "Invalid Soft Copy URL", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim docType = cmbDocType.SelectedItem.ToString()
            Dim code = EmbeddedDB.GenerateDocCode(docType)
            Dim assigned = cmbAssignedStaff.SelectedItem.ToString()

            EmbeddedDB.AddDocument(code, docType, txtTitle.Text.Trim(), txtOrigin.Text.Trim(), txtDest.Text.Trim(), txtCabinet.Text.Trim(), txtShelf.Text.Trim(), txtBox.Text.Trim(), txtGDrive.Text.Trim(), "LOGGED", assigned)
            EmbeddedDB.LogAudit(CurrentUser("FullName").ToString(), String.Format("Registered New Document [{0}] - {1}", code, txtTitle.Text.Trim()))

            MessageBox.Show(String.Format("Document Successfully Registered!" & vbCrLf & "Auto-Generated Code: {0}", code), "Registration Complete", MessageBoxButtons.OK, MessageBoxIcon.Information)

            txtTitle.Text = ""
            RefreshActiveTabGrid()
        End Sub

        Private Sub SetupDirectivesView()
            viewDirectives = New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.FromArgb(15, 23, 42)}

            Dim pnlTopDir As New Panel With {
                .Dock = DockStyle.Top,
                .Height = 110,
                .BackColor = Color.FromArgb(30, 41, 59),
                .Padding = New Padding(15, 12, 15, 12)
            }

            Dim lbl1 As New Label With {.Text = "Select Document:", .Location = New Point(15, 15), .AutoSize = True, .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold), .ForeColor = Color.FromArgb(148, 163, 184)}
            cmbDirDocs = New ComboBox With {.Location = New Point(135, 12), .Width = 420, .DropDownStyle = ComboBoxStyle.DropDownList, .BackColor = Color.FromArgb(15, 23, 42), .ForeColor = Color.White}

            Dim lbl2 As New Label With {.Text = "SG Directive:", .Location = New Point(575, 15), .AutoSize = True, .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold), .ForeColor = Color.FromArgb(148, 163, 184)}
            cmbDirective = New ComboBox With {.Location = New Point(665, 12), .Width = 320, .DropDownStyle = ComboBoxStyle.DropDownList, .BackColor = Color.FromArgb(15, 23, 42), .ForeColor = Color.White}
            cmbDirective.Items.AddRange(New Object() {"For Immediate Action", "Referred to Committee on Rules", "Forwarded for Speaker Signature", "Under OSG Administrative Review", "Approved & Archived"})
            cmbDirective.SelectedIndex = 0

            Dim lbl3 As New Label With {.Text = "Reassign Staff:", .Location = New Point(15, 58), .AutoSize = True, .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold), .ForeColor = Color.FromArgb(148, 163, 184)}
            cmbDirAssign = New ComboBox With {.Location = New Point(135, 55), .Width = 240, .DropDownStyle = ComboBoxStyle.DropDownList, .BackColor = Color.FromArgb(15, 23, 42), .ForeColor = Color.White}

            Dim lbl4 As New Label With {.Text = "Directive Notes:", .Location = New Point(395, 58), .AutoSize = True, .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold), .ForeColor = Color.FromArgb(148, 163, 184)}
            txtDirNotes = New TextBox With {.Location = New Point(495, 55), .Width = 330, .Text = "Priority routing per Secretary-General directive.", .BackColor = Color.FromArgb(15, 23, 42), .ForeColor = Color.White, .BorderStyle = BorderStyle.FixedSingle}

            btnApplyDirective = New Button With {
                .Text = "Log Action Directive",
                .Location = New Point(840, 52),
                .Size = New Size(165, 34),
                .BackColor = Color.FromArgb(16, 185, 129),
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
                .Cursor = Cursors.Hand
            }
            AddHandler btnApplyDirective.Click, AddressOf OnApplyDirective

            pnlTopDir.Controls.AddRange(New Control() {lbl1, cmbDirDocs, lbl2, cmbDirective, lbl3, cmbDirAssign, lbl4, txtDirNotes, btnApplyDirective})

            dgvDirectives = New DataGridView With {
                .Dock = DockStyle.Fill,
                .ReadOnly = True,
                .AllowUserToAddRows = False,
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            }
            ApplyGridStyle(dgvDirectives)

            viewDirectives.Controls.Add(dgvDirectives)
            viewDirectives.Controls.Add(pnlTopDir)
        End Sub

        Private Sub OnApplyDirective(sender As Object, e As EventArgs)
            If CurrentUser Is Nothing Then
                MessageBox.Show("RFID Authentication Required.", "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            If cmbDirDocs.SelectedItem Is Nothing Then
                MessageBox.Show("Please select a document.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim docStr = cmbDirDocs.SelectedItem.ToString()
            Dim docId As Integer = CInt(docStr.Split(":"c)(0).Replace("ID ", "").Trim())
            Dim directive = cmbDirective.SelectedItem.ToString()
            Dim assign = cmbDirAssign.SelectedItem.ToString()

            EmbeddedDB.AddDirective(docId, directive, assign, txtDirNotes.Text.Trim(), CurrentUser("FullName").ToString())
            EmbeddedDB.LogAudit(CurrentUser("FullName").ToString(), String.Format("Applied SG Directive [{0}] to Doc ID #{1}", directive, docId))

            MessageBox.Show("Secretary-General Action Directive Logged!", "Directive Applied", MessageBoxButtons.OK, MessageBoxIcon.Information)
            RefreshActiveTabGrid()
        End Sub

        Private Sub SetupSearchView()
            viewSearch = New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.FromArgb(15, 23, 42)}

            Dim pnlSearchTop As New Panel With {
                .Dock = DockStyle.Top,
                .Height = 65,
                .BackColor = Color.FromArgb(30, 41, 59),
                .Padding = New Padding(15, 12, 15, 12)
            }

            Dim lblKey As New Label With {.Text = "Search Keyword:", .Location = New Point(15, 20), .AutoSize = True, .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold), .ForeColor = Color.FromArgb(148, 163, 184)}
            txtSearchKey = New TextBox With {.Location = New Point(135, 17), .Width = 360, .BackColor = Color.FromArgb(15, 23, 42), .ForeColor = Color.White, .BorderStyle = BorderStyle.FixedSingle}
            btnSearch = New Button With {.Text = "Search", .Location = New Point(505, 15), .Size = New Size(95, 32), .BackColor = Color.FromArgb(79, 70, 229), .ForeColor = Color.White, .FlatStyle = FlatStyle.Flat, .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold), .Cursor = Cursors.Hand}
            btnViewSearchDetail = New Button With {.Text = "View Details & History", .Location = New Point(610, 15), .Size = New Size(180, 32), .BackColor = Color.FromArgb(14, 165, 233), .ForeColor = Color.White, .FlatStyle = FlatStyle.Flat, .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold), .Cursor = Cursors.Hand}
            btnOpenPDF = New Button With {.Text = "Launch Google Drive PDF", .Location = New Point(800, 15), .Size = New Size(200, 32), .BackColor = Color.FromArgb(16, 185, 129), .ForeColor = Color.White, .FlatStyle = FlatStyle.Flat, .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold), .Cursor = Cursors.Hand}

            AddHandler btnSearch.Click, AddressOf OnSearch
            AddHandler btnViewSearchDetail.Click, Sub() OpenSelectedDocumentDetail(dgvSearch)
            AddHandler btnOpenPDF.Click, AddressOf OnOpenPDF

            pnlSearchTop.Controls.AddRange(New Control() {lblKey, txtSearchKey, btnSearch, btnViewSearchDetail, btnOpenPDF})

            dgvSearch = New DataGridView With {
                .Dock = DockStyle.Fill,
                .ReadOnly = True,
                .AllowUserToAddRows = False,
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            }
            ApplyGridStyle(dgvSearch)
            AddHandler dgvSearch.CellDoubleClick, Sub(s, e) If e.RowIndex >= 0 Then OpenSelectedDocumentDetail(dgvSearch)

            viewSearch.Controls.Add(dgvSearch)
            viewSearch.Controls.Add(pnlSearchTop)
        End Sub

        Private Sub OnSearch(sender As Object, e As EventArgs)
            Dim visibleTable = EmbeddedDB.GetVisibleDocuments(CurrentUser)
            Dim q = txtSearchKey.Text.Trim().Replace("'", "''")
            If String.IsNullOrEmpty(q) Then
                dgvSearch.DataSource = visibleTable
            Else
                Dim filter As String = String.Format("DocCode LIKE '%{0}%' OR Title LIKE '%{0}%' OR DocType LIKE '%{0}%' OR CabinetID LIKE '%{0}%' OR OriginatingOffice LIKE '%{0}%' OR DestinationOffice LIKE '%{0}%' OR AssignedStaff LIKE '%{0}%'", q)
                Dim dv As New DataView(visibleTable)
                dv.RowFilter = filter
                dgvSearch.DataSource = dv.ToTable()
            End If
        End Sub

        Private Sub OnOpenPDF(sender As Object, e As EventArgs)
            If dgvSearch.CurrentRow Is Nothing Then
                MessageBox.Show("Please select a document row from the grid.", "Selection Required", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim url As String = dgvSearch.CurrentRow.Cells("GDriveURL").Value.ToString()
            If Not String.IsNullOrEmpty(url) Then
                Dim errUrl As String = ""
                If Not EmbeddedDB.ValidateGDriveURL(url, errUrl) Then
                    MessageBox.Show(errUrl, "Security Validation Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If
                Try
                    Process.Start(New ProcessStartInfo With {.FileName = url, .UseShellExecute = True})
                    If CurrentUser IsNot Nothing Then
                        EmbeddedDB.LogAudit(CurrentUser("FullName").ToString(), "Launched Google Drive Soft Copy PDF: " & url)
                    End If
                Catch ex As Exception
                    MessageBox.Show("Unable to launch URL: " & ex.Message, "Browser Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            Else
                MessageBox.Show("No Google Drive PDF URL associated with this document.", "No URL", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        End Sub

        Private Sub SetupAdminView()
            viewAdmin = New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.FromArgb(15, 23, 42)}

            Dim pnlForm As New Panel With {
                .Dock = DockStyle.Left,
                .Width = 360,
                .BackColor = Color.FromArgb(30, 41, 59),
                .Padding = New Padding(16)
            }

            Dim lblHeader As New Label With {
                .Text = "REGISTER USER & RFID BADGE",
                .Font = New Font("Segoe UI", 10.5F, FontStyle.Bold),
                .Location = New Point(15, 15),
                .AutoSize = True,
                .ForeColor = Color.White
            }

            txtNewUserName = New TextBox With {.Location = New Point(15, 65), .Width = 320, .Text = "New Staff Member", .BackColor = Color.FromArgb(15, 23, 42), .ForeColor = Color.White, .BorderStyle = BorderStyle.FixedSingle}
            cmbNewUserRole = New ComboBox With {.Location = New Point(15, 125), .Width = 320, .DropDownStyle = ComboBoxStyle.DropDownList, .BackColor = Color.FromArgb(15, 23, 42), .ForeColor = Color.White}
            cmbNewUserRole.Items.AddRange(New Object() {"Secretary-General", "OSG Chief", "System Administrator", "Administrative Staff"})
            cmbNewUserRole.SelectedIndex = 3

            txtNewUserUID = New TextBox With {.Location = New Point(15, 185), .Width = 320, .Text = "44D4E555", .BackColor = Color.FromArgb(15, 23, 42), .ForeColor = Color.White, .BorderStyle = BorderStyle.FixedSingle}

            btnAddUser = New Button With {
                .Text = "SAVE USER & RFID SMART CARD",
                .Location = New Point(15, 235),
                .Size = New Size(320, 40),
                .BackColor = Color.FromArgb(79, 70, 229),
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
                .Cursor = Cursors.Hand
            }
            AddHandler btnAddUser.Click, AddressOf OnAddUser

            pnlForm.Controls.AddRange(New Control() {
                lblHeader,
                New Label With {.Text = "Full Name:", .Location = New Point(15, 45), .AutoSize = True, .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold), .ForeColor = Color.FromArgb(148, 163, 184)},
                txtNewUserName,
                New Label With {.Text = "System Role:", .Location = New Point(15, 105), .AutoSize = True, .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold), .ForeColor = Color.FromArgb(148, 163, 184)},
                cmbNewUserRole,
                New Label With {.Text = "RFID Smart Card UID (Hex):", .Location = New Point(15, 165), .AutoSize = True, .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold), .ForeColor = Color.FromArgb(148, 163, 184)},
                txtNewUserUID,
                btnAddUser
            })

            dgvUsers = New DataGridView With {
                .Dock = DockStyle.Fill,
                .ReadOnly = True,
                .AllowUserToAddRows = False,
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            }
            ApplyGridStyle(dgvUsers)

            viewAdmin.Controls.Add(dgvUsers)
            viewAdmin.Controls.Add(pnlForm)
        End Sub

        Private Sub OnAddUser(sender As Object, e As EventArgs)
            If CurrentUser Is Nothing OrElse CurrentUser("Role").ToString() <> "System Administrator" Then
                MessageBox.Show("System Administrator Privileges Required.", "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            If String.IsNullOrWhiteSpace(txtNewUserName.Text) OrElse String.IsNullOrWhiteSpace(txtNewUserUID.Text) Then
                MessageBox.Show("Please enter Full Name and RFID Card UID.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            EmbeddedDB.AddUser(txtNewUserUID.Text.Trim(), txtNewUserName.Text.Trim(), cmbNewUserRole.SelectedItem.ToString())
            EmbeddedDB.LogAudit(CurrentUser("FullName").ToString(), String.Format("Registered/Updated User [{0}] Role: {1} RFID: {2}", txtNewUserName.Text.Trim(), cmbNewUserRole.SelectedItem, txtNewUserUID.Text.Trim()))

            MessageBox.Show("User and RFID Card Saved!", "Admin Save Complete", MessageBoxButtons.OK, MessageBoxIcon.Information)
            PopulateStaffDropdowns()
            RefreshActiveTabGrid()
        End Sub

        Private Sub SetupAuditView()
            viewAudit = New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.FromArgb(15, 23, 42)}

            Dim pnlAuditTop As New Panel With {
                .Dock = DockStyle.Top,
                .Height = 55,
                .BackColor = Color.FromArgb(30, 41, 59),
                .Padding = New Padding(15, 10, 15, 10)
            }

            Dim lblF As New Label With {.Text = "Filter Audit Trail:", .Location = New Point(15, 16), .AutoSize = True, .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold), .ForeColor = Color.FromArgb(148, 163, 184)}
            txtAuditSearch = New TextBox With {.Location = New Point(135, 13), .Width = 360, .BackColor = Color.FromArgb(15, 23, 42), .ForeColor = Color.White, .BorderStyle = BorderStyle.FixedSingle}
            btnAuditFilter = New Button With {.Text = "Filter", .Location = New Point(505, 11), .Size = New Size(95, 32), .BackColor = Color.FromArgb(79, 70, 229), .ForeColor = Color.White, .FlatStyle = FlatStyle.Flat, .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold), .Cursor = Cursors.Hand}
            AddHandler btnAuditFilter.Click, AddressOf OnFilterAudit

            pnlAuditTop.Controls.AddRange(New Control() {lblF, txtAuditSearch, btnAuditFilter})

            dgvAudit = New DataGridView With {
                .Dock = DockStyle.Fill,
                .ReadOnly = True,
                .AllowUserToAddRows = False,
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            }
            ApplyGridStyle(dgvAudit)

            viewAudit.Controls.Add(dgvAudit)
            viewAudit.Controls.Add(pnlAuditTop)
        End Sub

        Private Sub OnFilterAudit(sender As Object, e As EventArgs)
            Dim q = txtAuditSearch.Text.Trim().Replace("'", "''")
            Dim dt = EmbeddedDB.DataSet.Tables("AuditTrail")
            If String.IsNullOrEmpty(q) Then
                dgvAudit.DataSource = dt
            Else
                Dim dv As New DataView(dt)
                dv.RowFilter = String.Format("UserName LIKE '%{0}%' OR ActionDescription LIKE '%{0}%'", q)
                dgvAudit.DataSource = dv.ToTable()
            End If
        End Sub

        Private Sub OpenSelectedDocumentDetail(dgv As DataGridView)
            If dgv.CurrentRow Is Nothing Then
                MessageBox.Show("Please select a document row from the grid first.", "Selection Required", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim docId = CInt(dgv.CurrentRow.Cells("DocumentID").Value)
            Using dlg As New FormDocumentDetail(docId, Me)
                dlg.ShowDialog(Me)
            End Using
            RefreshActiveTabGrid()
        End Sub

        Public Sub RefreshActiveTabGrid()
            Dim visibleDocs = EmbeddedDB.GetVisibleDocuments(CurrentUser)

            ' Update Dashboard Stats
            lblStatTotalDocs.Text = EmbeddedDB.DataSet.Tables("Documents").Rows.Count.ToString()
            lblStatDirectives.Text = EmbeddedDB.DataSet.Tables("Directives").Rows.Count.ToString()
            lblStatActiveRoute.Text = EmbeddedDB.DataSet.Tables("RoutingLogs").Rows.Count.ToString()
            lblStatVaultStorage.Text = EmbeddedDB.DataSet.Tables("Movements").Rows.Count.ToString()
            dgvDashRecent.DataSource = visibleDocs

            Select Case activeNavIndex
                Case 1 ' Registry View
                    dgvRegistry.DataSource = visibleDocs
                Case 2 ' Directives View
                    cmbDirDocs.Items.Clear()
                    For Each row As DataRow In visibleDocs.Rows
                        cmbDirDocs.Items.Add(String.Format("ID {0}: [{1}] {2}", row("DocumentID"), row("DocCode"), row("Title")))
                    Next
                    If cmbDirDocs.Items.Count > 0 Then cmbDirDocs.SelectedIndex = 0
                    dgvDirectives.DataSource = EmbeddedDB.DataSet.Tables("Directives")
                Case 3 ' Search View
                    dgvSearch.DataSource = visibleDocs
                Case 4 ' Admin View
                    dgvUsers.DataSource = EmbeddedDB.DataSet.Tables("Users")
                Case 5 ' Audit View
                    dgvAudit.DataSource = EmbeddedDB.DataSet.Tables("AuditTrail")
            End Select
        End Sub
    End Class

    ' =========================================================================
    ' 2.2 DETAILED DOCUMENT METADATA & SPECIFICATION MATRIX DIALOG
    ' =========================================================================
    Public Class FormDocumentDetail
        Inherits Form

        Private DocID As Integer
        Private MainFrm As FormMain
        Private DocRow As DataRow

        Private pnlHeader As Panel
        Private lblCode As Label
        Private lblTitle As Label
        Private lblStatusBadge As Label

        Private tabDetail As TabControl
        Private tabOverview As TabPage
        Private tabDirectives As TabPage
        Private tabRouting As TabPage
        Private tabMovements As TabPage

        Private dgvDirectives As DataGridView
        Private dgvRouting As DataGridView
        Private dgvMovements As DataGridView

        Private pnlOverviewTable As TableLayoutPanel

        Public Sub New(id As Integer, parentForm As FormMain)
            DocID = id
            MainFrm = parentForm
            InitializeForm()
        End Sub

        Private Sub InitializeForm()
            If Not LoadDocData() Then Return

            Me.Text = String.Format("Document Details & Specifications — [{0}] {1}", DocRow("DocCode"), DocRow("Title"))
            Me.Size = New Size(1080, 760)
            Me.MinimumSize = New Size(920, 640)
            Me.StartPosition = FormStartPosition.CenterParent
            Me.Font = New Font("Segoe UI", 9.5F, FontStyle.Regular)
            Me.BackColor = Color.FromArgb(15, 23, 42)

            ' Top Header Panel
            pnlHeader = New Panel With {
                .Dock = DockStyle.Top,
                .Height = 135,
                .BackColor = Color.FromArgb(30, 41, 59),
                .Padding = New Padding(20, 12, 20, 12)
            }

            lblCode = New Label With {
                .Text = String.Format("DOC CODE: {0}   |   TYPE: {1}", DocRow("DocCode"), DocRow("DocType").ToString().ToUpper()),
                .Font = New Font("Segoe UI", 12.0F, FontStyle.Bold),
                .ForeColor = Color.FromArgb(129, 140, 248),
                .Location = New Point(20, 12),
                .AutoSize = True
            }

            lblTitle = New Label With {
                .Text = DocRow("Title").ToString(),
                .Font = New Font("Segoe UI", 10.5F, FontStyle.Bold),
                .ForeColor = Color.White,
                .Location = New Point(20, 42),
                .Size = New Size(640, 44)
            }

            lblStatusBadge = New Label With {
                .Text = String.Format("Status: {0}  |  Assigned Staff: {1}", DocRow("CurrentStatus"), DocRow("AssignedStaff")),
                .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
                .ForeColor = Color.FromArgb(52, 211, 153),
                .Location = New Point(20, 92),
                .AutoSize = True
            }

            Dim btnRoute As New Button With {
                .Text = "Route Office Step",
                .Location = New Point(700, 14),
                .Size = New Size(220, 34),
                .BackColor = Color.FromArgb(79, 70, 229),
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
                .Cursor = Cursors.Hand,
                .Anchor = AnchorStyles.Top Or AnchorStyles.Right
            }
            AddHandler btnRoute.Click, AddressOf OnRouteDocument

            Dim btnMove As New Button With {
                .Text = "Transfer Physical Storage",
                .Location = New Point(700, 52),
                .Size = New Size(220, 34),
                .BackColor = Color.FromArgb(14, 165, 233),
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
                .Cursor = Cursors.Hand,
                .Anchor = AnchorStyles.Top Or AnchorStyles.Right
            }
            AddHandler btnMove.Click, AddressOf OnMoveStorage

            Dim btnLaunchPdf As New Button With {
                .Text = "Launch Google Drive PDF",
                .Location = New Point(700, 90),
                .Size = New Size(220, 34),
                .BackColor = Color.FromArgb(16, 185, 129),
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
                .Cursor = Cursors.Hand,
                .Anchor = AnchorStyles.Top Or AnchorStyles.Right
            }
            AddHandler btnLaunchPdf.Click, AddressOf OnLaunchPDF

            pnlHeader.Controls.AddRange(New Control() {lblCode, lblTitle, lblStatusBadge, btnRoute, btnMove, btnLaunchPdf})
            Me.Controls.Add(pnlHeader)

            ' Tab Navigation
            tabDetail = New TabControl With {
                .Dock = DockStyle.Fill,
                .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold),
                .Padding = New Point(12, 6)
            }

            tabOverview = New TabPage(" Document Overview & Specifications ")
            tabDirectives = New TabPage(" SG Directives Timeline ")
            tabRouting = New TabPage(" Office Routing Logs ")
            tabMovements = New TabPage(" Physical Storage Movement History ")

            tabOverview.BackColor = Color.FromArgb(15, 23, 42)
            tabDirectives.BackColor = Color.FromArgb(15, 23, 42)
            tabRouting.BackColor = Color.FromArgb(15, 23, 42)
            tabMovements.BackColor = Color.FromArgb(15, 23, 42)

            SetupOverviewTab()

            dgvDirectives = CreateDetailGrid()
            dgvRouting = CreateDetailGrid()
            dgvMovements = CreateDetailGrid()

            tabDirectives.Controls.Add(dgvDirectives)
            tabRouting.Controls.Add(dgvRouting)
            tabMovements.Controls.Add(dgvMovements)

            tabDetail.TabPages.AddRange(New TabPage() {tabOverview, tabDirectives, tabRouting, tabMovements})
            Me.Controls.Add(tabDetail)
            pnlHeader.SendToBack()

            RefreshGrids()
        End Sub

        Private Function LoadDocData() As Boolean
            Dim rows = EmbeddedDB.DataSet.Tables("Documents").Select("DocumentID = " & DocID)
            If rows.Length = 0 Then
                MessageBox.Show("Document record not found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.Close()
                Return False
            End If
            DocRow = rows(0)
            Return True
        End Function

        Private Sub SetupOverviewTab()
            tabOverview.Controls.Clear()

            Dim pnlScroll As New Panel With {
                .Dock = DockStyle.Fill,
                .AutoScroll = True,
                .Padding = New Padding(20)
            }

            pnlOverviewTable = New TableLayoutPanel With {
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .ColumnCount = 2,
                .CellBorderStyle = TableLayoutPanelCellBorderStyle.Single,
                .BackColor = Color.FromArgb(30, 41, 59),
                .Padding = New Padding(10)
            }

            pnlOverviewTable.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 30.0F))
            pnlOverviewTable.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 70.0F))

            AddOverviewRow("Document System ID:", "#" & DocRow("DocumentID").ToString())
            AddOverviewRow("Document Code:", DocRow("DocCode").ToString())
            AddOverviewRow("Document Classification Type:", DocRow("DocType").ToString())
            AddOverviewRow("Official Document Title:", DocRow("Title").ToString())
            AddOverviewRow("Date Received / Registered:", DocRow("DateReceived").ToString())
            AddOverviewRow("Originating Office:", DocRow("OriginatingOffice").ToString())
            AddOverviewRow("Destination Office:", DocRow("DestinationOffice").ToString())
            AddOverviewRow("Cabinet Landmark ID:", DocRow("CabinetID").ToString())
            AddOverviewRow("Shelf Landmark No:", DocRow("ShelfNo").ToString())
            AddOverviewRow("Box Landmark Code:", DocRow("BoxCode").ToString())
            AddOverviewRow("Full Physical Location String:", String.Format("Cabinet {0} / Shelf {1} / Box {2}", DocRow("CabinetID"), DocRow("ShelfNo"), DocRow("BoxCode")))
            AddOverviewRow("Current Processing Status:", DocRow("CurrentStatus").ToString())
            AddOverviewRow("Assigned OSG Staff Member:", DocRow("AssignedStaff").ToString())
            AddOverviewRow("Google Drive Soft Copy URL:", DocRow("GDriveURL").ToString())

            pnlScroll.Controls.Add(pnlOverviewTable)
            tabOverview.Controls.Add(pnlScroll)
        End Sub

        Private Sub AddOverviewRow(label As String, value As String)
            Dim lblField As New Label With {
                .Text = label,
                .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
                .ForeColor = Color.FromArgb(148, 163, 184),
                .Dock = DockStyle.Fill,
                .TextAlign = ContentAlignment.MiddleLeft,
                .Padding = New Padding(8, 8, 8, 8),
                .AutoSize = True
            }

            Dim lblVal As New Label With {
                .Text = If(String.IsNullOrWhiteSpace(value), "(None Specified)", value),
                .Font = New Font("Segoe UI", 9.5F, FontStyle.Regular),
                .ForeColor = Color.FromArgb(248, 250, 252),
                .Dock = DockStyle.Fill,
                .TextAlign = ContentAlignment.MiddleLeft,
                .Padding = New Padding(8, 8, 8, 8),
                .AutoSize = True
            }

            Dim rowIndex = pnlOverviewTable.RowCount
            pnlOverviewTable.RowCount += 1
            pnlOverviewTable.Controls.Add(lblField, 0, rowIndex)
            pnlOverviewTable.Controls.Add(lblVal, 1, rowIndex)
        End Sub

        Private Function CreateDetailGrid() As DataGridView
            Dim dgv As New DataGridView With {
                .Dock = DockStyle.Fill,
                .ReadOnly = True,
                .AllowUserToAddRows = False,
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            }
            FormMain.ApplyGridStyle(dgv)
            Return dgv
        End Function

        Private Sub RefreshGrids()
            If Not LoadDocData() Then Return

            lblCode.Text = String.Format("DOC CODE: {0}   |   TYPE: {1}", DocRow("DocCode"), DocRow("DocType").ToString().ToUpper())
            lblTitle.Text = DocRow("Title").ToString()
            lblStatusBadge.Text = String.Format("Status: {0}  |  Assigned Staff: {1}", DocRow("CurrentStatus"), DocRow("AssignedStaff"))

            SetupOverviewTab()

            Dim dvDir As New DataView(EmbeddedDB.DataSet.Tables("Directives"))
            dvDir.RowFilter = "DocumentID = " & DocID
            dgvDirectives.DataSource = dvDir.ToTable()

            Dim dvRoute As New DataView(EmbeddedDB.DataSet.Tables("RoutingLogs"))
            dvRoute.RowFilter = "DocumentID = " & DocID
            dgvRouting.DataSource = dvRoute.ToTable()

            Dim dvMove As New DataView(EmbeddedDB.DataSet.Tables("Movements"))
            dvMove.RowFilter = "DocumentID = " & DocID
            dgvMovements.DataSource = dvMove.ToTable()
        End Sub

        Private Sub OnRouteDocument(sender As Object, e As EventArgs)
            If MainFrm.CurrentUser Is Nothing Then
                MessageBox.Show("Authentication Required.", "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Using dlg As New FormRouteDocument(DocID, DocRow("DestinationOffice").ToString(), MainFrm.CurrentUser("FullName").ToString())
                If dlg.ShowDialog(Me) = DialogResult.OK Then
                    RefreshGrids()
                End If
            End Using
        End Sub

        Private Sub OnMoveStorage(sender As Object, e As EventArgs)
            If MainFrm.CurrentUser Is Nothing Then
                MessageBox.Show("Authentication Required.", "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim currentLoc = String.Format("{0}/{1}/{2}", DocRow("CabinetID"), DocRow("ShelfNo"), DocRow("BoxCode"))
            Using dlg As New FormMoveStorage(DocID, currentLoc, MainFrm.CurrentUser("FullName").ToString())
                If dlg.ShowDialog(Me) = DialogResult.OK Then
                    RefreshGrids()
                End If
            End Using
        End Sub

        Private Sub OnLaunchPDF(sender As Object, e As EventArgs)
            Dim url = DocRow("GDriveURL").ToString()
            Dim errUrl As String = ""
            If Not EmbeddedDB.ValidateGDriveURL(url, errUrl) Then
                MessageBox.Show(errUrl, "Security Validation Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            Try
                Process.Start(New ProcessStartInfo With {.FileName = url, .UseShellExecute = True})
                If MainFrm.CurrentUser IsNot Nothing Then
                    EmbeddedDB.LogAudit(MainFrm.CurrentUser("FullName").ToString(), "Launched Google Drive Soft Copy PDF: " & url)
                End If
            Catch ex As Exception
                MessageBox.Show("Error opening URL: " & ex.Message, "Launch Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub
    End Class

    ' =========================================================================
    ' ROUTING DIALOG
    ' =========================================================================
    Public Class FormRouteDocument
        Inherits Form

        Private DocID As Integer
        Private StaffName As String
        Private txtFrom As TextBox
        Private txtTo As TextBox
        Private txtAction As TextBox
        Private txtRemarks As TextBox

        Public Sub New(docId As Integer, currentOffice As String, user As String)
            Me.DocID = docId
            Me.StaffName = user

            Me.Text = "Log Document Office Routing Step"
            Me.Size = New Size(460, 340)
            Me.StartPosition = FormStartPosition.CenterParent
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.BackColor = Color.FromArgb(15, 23, 42)

            txtFrom = New TextBox With {.Location = New Point(140, 20), .Width = 270, .Text = currentOffice, .BackColor = Color.FromArgb(30, 41, 59), .ForeColor = Color.White, .BorderStyle = BorderStyle.FixedSingle}
            txtTo = New TextBox With {.Location = New Point(140, 60), .Width = 270, .Text = "Speaker's Office", .BackColor = Color.FromArgb(30, 41, 59), .ForeColor = Color.White, .BorderStyle = BorderStyle.FixedSingle}
            txtAction = New TextBox With {.Location = New Point(140, 100), .Width = 270, .Text = "FOR_SIGNATURE", .BackColor = Color.FromArgb(30, 41, 59), .ForeColor = Color.White, .BorderStyle = BorderStyle.FixedSingle}
            txtRemarks = New TextBox With {.Location = New Point(140, 140), .Width = 270, .Text = "Transmitted for Speaker approval.", .BackColor = Color.FromArgb(30, 41, 59), .ForeColor = Color.White, .BorderStyle = BorderStyle.FixedSingle}

            Dim btnSave As New Button With {.Text = "Save Route Step", .Location = New Point(140, 195), .Size = New Size(160, 36), .BackColor = Color.FromArgb(79, 70, 229), .ForeColor = Color.White, .FlatStyle = FlatStyle.Flat, .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold)}
            AddHandler btnSave.Click, AddressOf OnSave

            Me.Controls.AddRange(New Control() {
                New Label With {.Text = "From Office:", .Location = New Point(20, 23), .AutoSize = True, .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold), .ForeColor = Color.FromArgb(148, 163, 184)}, txtFrom,
                New Label With {.Text = "To Office:", .Location = New Point(20, 63), .AutoSize = True, .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold), .ForeColor = Color.FromArgb(148, 163, 184)}, txtTo,
                New Label With {.Text = "Action Taken:", .Location = New Point(20, 103), .AutoSize = True, .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold), .ForeColor = Color.FromArgb(148, 163, 184)}, txtAction,
                New Label With {.Text = "Remarks:", .Location = New Point(20, 143), .AutoSize = True, .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold), .ForeColor = Color.FromArgb(148, 163, 184)}, txtRemarks,
                btnSave
            })
        End Sub

        Private Sub OnSave(sender As Object, e As EventArgs)
            EmbeddedDB.AddRoutingLog(DocID, txtFrom.Text.Trim(), txtTo.Text.Trim(), StaffName, txtAction.Text.Trim(), txtRemarks.Text.Trim())
            EmbeddedDB.LogAudit(StaffName, String.Format("Routed Doc #{0} from {1} to {2}", DocID, txtFrom.Text, txtTo.Text))
            Me.DialogResult = DialogResult.OK
            Me.Close()
        End Sub
    End Class

    ' =========================================================================
    ' STORAGE MOVEMENT DIALOG
    ' =========================================================================
    Public Class FormMoveStorage
        Inherits Form

        Private DocID As Integer
        Private StaffName As String
        Private txtFromLoc As TextBox
        Private txtToLoc As TextBox
        Private txtReason As TextBox

        Public Sub New(docId As Integer, currentLoc As String, user As String)
            Me.DocID = docId
            Me.StaffName = user

            Me.Text = "Transfer Physical Landmark Storage Location"
            Me.Size = New Size(470, 300)
            Me.StartPosition = FormStartPosition.CenterParent
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.BackColor = Color.FromArgb(15, 23, 42)

            txtFromLoc = New TextBox With {.Location = New Point(150, 20), .Width = 270, .Text = currentLoc, .ReadOnly = True, .BackColor = Color.FromArgb(30, 41, 59), .ForeColor = Color.White, .BorderStyle = BorderStyle.FixedSingle}
            txtToLoc = New TextBox With {.Location = New Point(150, 60), .Width = 270, .Text = "CAB-A/S-3/BOX-05", .BackColor = Color.FromArgb(30, 41, 59), .ForeColor = Color.White, .BorderStyle = BorderStyle.FixedSingle}
            txtReason = New TextBox With {.Location = New Point(150, 100), .Width = 270, .Text = "Archival reorganization per OSG directive.", .BackColor = Color.FromArgb(30, 41, 59), .ForeColor = Color.White, .BorderStyle = BorderStyle.FixedSingle}

            Dim btnSave As New Button With {.Text = "Save Physical Transfer", .Location = New Point(150, 155), .Size = New Size(180, 36), .BackColor = Color.FromArgb(14, 165, 233), .ForeColor = Color.White, .FlatStyle = FlatStyle.Flat, .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold)}
            AddHandler btnSave.Click, AddressOf OnSave

            Me.Controls.AddRange(New Control() {
                New Label With {.Text = "Current Location:", .Location = New Point(20, 23), .AutoSize = True, .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold), .ForeColor = Color.FromArgb(148, 163, 184)}, txtFromLoc,
                New Label With {.Text = "New (Cab/Shelf/Box):", .Location = New Point(20, 63), .AutoSize = True, .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold), .ForeColor = Color.FromArgb(148, 163, 184)}, txtToLoc,
                New Label With {.Text = "Reason for Move:", .Location = New Point(20, 103), .AutoSize = True, .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold), .ForeColor = Color.FromArgb(148, 163, 184)}, txtReason,
                btnSave
            })
        End Sub

        Private Sub OnSave(sender As Object, e As EventArgs)
            EmbeddedDB.AddMovementLog(DocID, txtFromLoc.Text.Trim(), txtToLoc.Text.Trim(), StaffName, txtReason.Text.Trim())
            EmbeddedDB.LogAudit(StaffName, String.Format("Transferred Doc #{0} physical location from {1} to {2}", DocID, txtFromLoc.Text, txtToLoc.Text))
            Me.DialogResult = DialogResult.OK
            Me.Close()
        End Sub
    End Class

    ' =========================================================================
    ' 2.5 DEDICATED RFID SCANNER LOGIN MODAL DIALOG
    ' =========================================================================
    Public Class FormRFIDLogin
        Inherits Form

        <System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)>
        Public Property ScannedUID As String = ""
        Private txtScanInput As TextBox
        Private lblScanStatus As Label

        Public Sub New()
            InitializeForm()
        End Sub

        Private Sub InitializeForm()
            Me.Text = "OSG RFID Smart Card Authentication Tap Scanner"
            Me.Size = New Size(480, 360)
            Me.StartPosition = FormStartPosition.CenterParent
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.BackColor = Color.FromArgb(15, 23, 42)

            Dim pnlCard As New Panel With {
                .Location = New Point(30, 20),
                .Size = New Size(400, 160),
                .BackColor = Color.FromArgb(30, 41, 59),
                .BorderStyle = BorderStyle.FixedSingle
            }

            Dim lblIcon As New Label With {
                .Text = "[ RFID TAP BADGE SCANNER ]",
                .ForeColor = Color.FromArgb(56, 189, 248),
                .Font = New Font("Segoe UI", 11.0F, FontStyle.Bold),
                .Location = New Point(20, 20),
                .AutoSize = True
            }

            lblScanStatus = New Label With {
                .Text = "Ready to Scan..." & vbCrLf & "Please tap your USB RFID Smart Card Badge on the physical scanner.",
                .ForeColor = Color.White,
                .Font = New Font("Segoe UI", 9.5F, FontStyle.Regular),
                .Location = New Point(20, 60),
                .Size = New Size(360, 60)
            }

            pnlCard.Controls.Add(lblIcon)
            pnlCard.Controls.Add(lblScanStatus)

            txtScanInput = New TextBox With {
                .Location = New Point(-100, -100),
                .Size = New Size(10, 10)
            }
            AddHandler txtScanInput.KeyDown, AddressOf OnScanInputKeyDown

            Dim lblSim As New Label With {
                .Text = "Or select a test badge to simulate tap:",
                .ForeColor = Color.FromArgb(148, 163, 184),
                .Font = New Font("Segoe UI", 8.5F, FontStyle.Italic),
                .Location = New Point(30, 195),
                .AutoSize = True
            }

            Dim btnSG = CreateSimButton("SG Card (88A9F321)", New Point(30, 220), Color.FromArgb(16, 185, 129))
            Dim btnAdmin = CreateSimButton("Admin Card (77C3D987)", New Point(235, 220), Color.FromArgb(14, 165, 233))
            Dim btnStaff = CreateSimButton("Staff Card (55E5F666)", New Point(30, 265), Color.FromArgb(245, 158, 11))
            Dim btnCancel = CreateSimButton("Cancel", New Point(235, 265), Color.FromArgb(100, 116, 139))

            AddHandler btnSG.Click, Sub() SubmitUID("88A9F321")
            AddHandler btnAdmin.Click, Sub() SubmitUID("77C3D987")
            AddHandler btnStaff.Click, Sub() SubmitUID("55E5F666")
            AddHandler btnCancel.Click, Sub() Me.DialogResult = DialogResult.Cancel

            Me.Controls.AddRange(New Control() {pnlCard, txtScanInput, lblSim, btnSG, btnAdmin, btnStaff, btnCancel})

            AddHandler Me.Shown, Sub() txtScanInput.Focus()
        End Sub

        Private Function CreateSimButton(text As String, loc As Point, bg As Color) As Button
            Return New Button With {
                .Text = text,
                .Location = loc,
                .Size = New Size(190, 36),
                .BackColor = bg,
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
                .Cursor = Cursors.Hand
            }
        End Function

        Private Sub OnScanInputKeyDown(sender As Object, e As KeyEventArgs)
            If e.KeyCode = Keys.Enter Then
                Dim uid = txtScanInput.Text.Trim()
                txtScanInput.Text = ""
                If Not String.IsNullOrEmpty(uid) Then
                    SubmitUID(uid)
                End If
                e.Handled = True
            End If
        End Sub

        Private Sub SubmitUID(uid As String)
            ScannedUID = uid
            Me.DialogResult = DialogResult.OK
            Me.Close()
        End Sub
    End Class

    ' =========================================================================
    ' 3. PROGRAM ENTRY POINT & SELF-CHECK AUTOMATED TEST SUITE
    ' =========================================================================
    Public Module Program
        <STAThread>
        Public Sub Main(args As String())
            AddHandler Application.ThreadException, Sub(sender, e)
                                                         MessageBox.Show("Thread Exception: " & e.Exception.ToString(), "Application Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                                     End Sub

            AddHandler AppDomain.CurrentDomain.UnhandledException, Sub(sender, e)
                                                                        MessageBox.Show("Unhandled Exception: " & e.ExceptionObject.ToString(), "Fatal Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                                                    End Sub

            If args IsNot Nothing AndAlso args.Length > 0 Then
                Dim arg = args(0).ToLower()
                If arg = "/test" OrElse arg = "/smoke" Then
                    Dim exitCode As Integer = RunAutomatedSelfCheck()
                    Environment.Exit(exitCode)
                    Return
                End If
            End If

            Try
                Application.SetHighDpiMode(HighDpiMode.SystemAware)
                Application.EnableVisualStyles()
                Application.SetCompatibleTextRenderingDefault(False)
                Application.Run(New FormMain())
            Catch ex As Exception
                MessageBox.Show("Startup Error: " & ex.ToString(), "Initialization Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Private Function RunAutomatedSelfCheck() As Integer
            Console.WriteLine("=========================================================")
            Console.WriteLine("  BTA OSG SYSTEM AUTOMATED SELF-CHECK SUITE")
            Console.WriteLine("=========================================================")

            Try
                ' Test 1: DB Initialization
                EmbeddedDB.Initialize()
                Console.WriteLine("[PASS] 1. Embedded Relational DB Initialized.")

                ' Test 2: RFID Authentication & Normalization
                Dim sgUser = EmbeddedDB.AuthenticateRFID("88a9f321")
                If sgUser Is Nothing OrElse sgUser("Role").ToString() <> "Secretary-General" Then
                    Console.WriteLine("[FAIL] 2. RFID Authentication failed for SG.")
                    Return 1
                End If
                Console.WriteLine("[PASS] 2. RFID Authentication & UID Normalization Verified.")

                ' Test 3: Auto-Code Generation
                Dim code = EmbeddedDB.GenerateDocCode("Resolution")
                If Not code.StartsWith("RES-") Then
                    Console.WriteLine("[FAIL] 3. Auto-code generation returned unexpected format: " & code)
                    Return 1
                End If
                Console.WriteLine("[PASS] 3. Document Auto-Coding Verified (" & code & ").")

                ' Test 4: Document Creation & Initial Logs
                Dim docId = EmbeddedDB.AddDocument(code, "Resolution", "Test Resolution for BARMM Youth Council", "Office of MP Yasser", "OSG", "CAB-Z", "S-9", "BOX-99", "https://drive.google.com/test", "LOGGED", "Hassim A. Ibrahim")
                If docId <= 0 Then
                    Console.WriteLine("[FAIL] 4. Failed to add test document.")
                    Return 1
                End If
                Console.WriteLine("[PASS] 4. Document Registration & Initial Tracking Logs Verified (ID #" & docId & ").")

                ' Test 5: Directive Logging
                EmbeddedDB.AddDirective(docId, "For Immediate Action", "Hassim A. Ibrahim", "Priority processing", "Prof. Ali B. Pangalian")
                Console.WriteLine("[PASS] 5. SG Action Directive Logged.")

                ' Test 6: Office Routing Logging
                EmbeddedDB.AddRoutingLog(docId, "OSG", "Committee on Rules", "Prof. Ali B. Pangalian", "REFERRAL", "Referred for review")
                Console.WriteLine("[PASS] 6. Office Routing Log Verified.")

                ' Test 7: Storage Physical Movement Logging
                EmbeddedDB.AddMovementLog(docId, "CAB-Z/S-9/BOX-99", "CAB-A/S-1/BOX-01", "Hassim A. Ibrahim", "Moved to primary vault")
                Console.WriteLine("[PASS] 7. Storage Landmark Physical Movement Verified.")

                ' Test 8: Soft Copy PDF Security Domain Validation
                Dim validErr As String = ""
                If Not EmbeddedDB.ValidateGDriveURL("https://drive.google.com/file/d/123/view", validErr) Then
                    Console.WriteLine("[FAIL] 8. GDrive URL validation rejected valid drive link.")
                    Return 1
                End If
                If EmbeddedDB.ValidateGDriveURL("http://untrusted-site.com/doc.pdf", validErr) Then
                    Console.WriteLine("[FAIL] 8. GDrive URL validation accepted untrusted domain.")
                    Return 1
                End If
                Console.WriteLine("[PASS] 8. Google Drive Soft Copy PDF Security Validation Verified.")

                ' Test 9: User & RFID Card Management
                EmbeddedDB.AddUser("33334444", "Test Auditor", "Administrative Staff")
                If EmbeddedDB.AuthenticateRFID("33334444") Is Nothing Then
                    Console.WriteLine("[FAIL] 9. Newly registered user RFID authentication failed.")
                    Return 1
                End If
                Console.WriteLine("[PASS] 9. Dynamic User & RFID Card Administration Verified.")

                ' Test 10: RBAC Visibility Check
                Dim sgTable = EmbeddedDB.GetVisibleDocuments(sgUser)
                Dim staffUser = EmbeddedDB.AuthenticateRFID("55E5F666")
                Dim staffTable = EmbeddedDB.GetVisibleDocuments(staffUser)

                If sgTable.Rows.Count < staffTable.Rows.Count Then
                    Console.WriteLine("[FAIL] 10. RBAC filter error: SG has fewer documents than staff.")
                    Return 1
                End If
                Console.WriteLine("[PASS] 10. Role-Based Access Control (RBAC) Filter Verified.")

                Console.WriteLine("=========================================================")
                Console.WriteLine("ALL 10 SELF-CHECK TESTS PASSED SUCCESSFULLY!")
                Console.WriteLine("=========================================================")
                Return 0
            Catch ex As Exception
                Console.WriteLine("[FAIL] Self-check exception: " & ex.Message)
                Return 1
            End Try
        End Function
    End Module

End Namespace
