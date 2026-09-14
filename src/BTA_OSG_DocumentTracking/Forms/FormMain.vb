Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Diagnostics
Imports System.Drawing
Imports System.Windows.Forms

Namespace BTA_OSG
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
            Dim btn As New Button With {
                .Text = text,
                .Location = loc,
                .Size = New Size(width, 36),
                .BackColor = bg,
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
                .Cursor = Cursors.Hand
            }
            btn.FlatAppearance.BorderSize = 0
            Return btn
        End Function

        Private Sub ShowRFIDLoginDialog()
            Using dlg As New FormLogin()
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
            btnRegister.FlatAppearance.BorderSize = 0
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
            btnViewRegistryDetail.FlatAppearance.BorderSize = 0
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
            Dim assigned = If(cmbAssignedStaff.SelectedItem IsNot Nothing, cmbAssignedStaff.SelectedItem.ToString(), "")

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
            btnApplyDirective.FlatAppearance.BorderSize = 0
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
            Dim assign = If(cmbDirAssign.SelectedItem IsNot Nothing, cmbDirAssign.SelectedItem.ToString(), "")

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
            btnSearch.FlatAppearance.BorderSize = 0
            btnViewSearchDetail = New Button With {.Text = "View Details & History", .Location = New Point(610, 15), .Size = New Size(180, 32), .BackColor = Color.FromArgb(14, 165, 233), .ForeColor = Color.White, .FlatStyle = FlatStyle.Flat, .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold), .Cursor = Cursors.Hand}
            btnViewSearchDetail.FlatAppearance.BorderSize = 0
            btnOpenPDF = New Button With {.Text = "Launch Google Drive PDF", .Location = New Point(800, 15), .Size = New Size(200, 32), .BackColor = Color.FromArgb(16, 185, 129), .ForeColor = Color.White, .FlatStyle = FlatStyle.Flat, .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold), .Cursor = Cursors.Hand}
            btnOpenPDF.FlatAppearance.BorderSize = 0

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
            btnAddUser.FlatAppearance.BorderSize = 0
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
            btnAuditFilter.FlatAppearance.BorderSize = 0
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

        Public Sub OpenSelectedDocumentDetail(dgv As DataGridView)
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
End Namespace
