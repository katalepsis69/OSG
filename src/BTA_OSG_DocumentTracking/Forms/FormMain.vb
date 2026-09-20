Option Explicit On
Option Strict On

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
        Public CurrentSession As SessionContext = Nothing

        ' Layout Panels
        Private pnlSidebar As Panel
        Private pnlHeader As Panel
        Private pnlContent As Panel
        Private stsFooter As StatusStrip
        Private lblStatusMessage As ToolStripStatusLabel
        Private lblStatusCount As ToolStripStatusLabel
        Private lblStatusClock As ToolStripStatusLabel
        Private epValidation As ErrorProvider

        ' Header Controls
        Private lblTitle As Label
        Private lblUserBadge As Label
        Private btnScanRFID As Button
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
        Private lblDashWatermark As Label

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
        Private lblRegistryWatermark As Label
        Private btnViewRegistryDetail As Button

        ' Directives Controls
        Private cmbDirDocs As ComboBox
        Private cmbDirective As ComboBox
        Private cmbDirAssign As ComboBox
        Private txtDirNotes As TextBox
        Private btnApplyDirective As Button
        Private dgvDirectives As DataGridView
        Private lblDirectivesWatermark As Label

        ' Search Controls
        Private txtSearchKey As TextBox
        Private btnSearch As Button
        Private btnClearSearch As Button
        Private dgvSearch As DataGridView
        Private lblSearchWatermark As Label
        Private btnViewSearchDetail As Button
        Private btnOpenPDF As Button

        ' User Admin Controls
        Private dgvUsers As DataGridView
        Private lblUsersWatermark As Label
        Private txtNewUserName As TextBox
        Private cmbNewUserRole As ComboBox
        Private txtNewUserUID As TextBox
        Private btnAddUser As Button

        ' Audit Controls
        Private txtAuditSearch As TextBox
        Private btnAuditFilter As Button
        Private btnAuditReset As Button
        Private dgvAudit As DataGridView
        Private lblAuditWatermark As Label

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
            Me.Font = CivicCalmTheme.FontBody
            Me.BackColor = CivicCalmTheme.ColorCanvas

            epValidation = New ErrorProvider With {
                .BlinkStyle = ErrorBlinkStyle.NeverBlink
            }

            SetupFooterStatusStrip()
            SetupSidebar()
            SetupHeader()

            pnlContent = New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = CivicCalmTheme.ColorCanvas,
                .Padding = New Padding(16)
            }
            Me.Controls.Add(pnlContent)
            pnlContent.BringToFront()

            SetupDashboardView()
            SetupRegistryView()
            SetupDirectivesView()
            SetupSearchView()
            SetupAdminView()
            SetupAuditView()

            SwitchNavView(0)
        End Sub

        Private Sub SetupFooterStatusStrip()
            stsFooter = New StatusStrip With {
                .Dock = DockStyle.Bottom,
                .BackColor = CivicCalmTheme.ColorWell,
                .ForeColor = CivicCalmTheme.ColorInk,
                .Font = CivicCalmTheme.FontMicrocopy,
                .SizingGrip = False
            }

            lblStatusMessage = New ToolStripStatusLabel With {
                .Text = "Ready",
                .Spring = True,
                .TextAlign = ContentAlignment.MiddleLeft,
                .ForeColor = CivicCalmTheme.ColorInk
            }

            lblStatusCount = New ToolStripStatusLabel With {
                .Text = "0 Records",
                .BorderSides = ToolStripStatusLabelBorderSides.Left,
                .BorderStyle = Border3DStyle.Etched,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .Padding = New Padding(8, 0, 8, 0)
            }

            lblStatusClock = New ToolStripStatusLabel With {
                .Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                .BorderSides = ToolStripStatusLabelBorderSides.Left,
                .BorderStyle = Border3DStyle.Etched,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .Padding = New Padding(8, 0, 8, 0)
            }

            stsFooter.Items.AddRange(New ToolStripItem() {lblStatusMessage, lblStatusCount, lblStatusClock})
            Me.Controls.Add(stsFooter)
        End Sub

        Private Sub SetupSidebar()
            pnlSidebar = New Panel With {
                .Dock = DockStyle.Left,
                .Width = 240,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(12)
            }

            Dim pnlBrand As New Panel With {
                .Dock = DockStyle.Top,
                .Height = 85,
                .BackColor = CivicCalmTheme.ColorSurface
            }
            Dim lblBrand As New Label With {
                .Text = "BTA PARLIAMENT" & vbCrLf & "OFFICE OF THE SG",
                .Font = CivicCalmTheme.FontFormTitle,
                .ForeColor = CivicCalmTheme.ColorPrimary,
                .Dock = DockStyle.Top,
                .Height = 50,
                .TextAlign = ContentAlignment.MiddleLeft
            }
            Dim lblBrandSub As New Label With {
                .Text = "DOCUMENT TRACKING SYSTEM",
                .Font = CivicCalmTheme.FontMicrocopy,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .Dock = DockStyle.Top,
                .Height = 25,
                .TextAlign = ContentAlignment.MiddleLeft
            }
            pnlBrand.Controls.AddRange(New Control() {lblBrandSub, lblBrand})

            Dim pnlNavStack As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.TopDown,
                .WrapContents = False,
                .AutoScroll = True,
                .BackColor = CivicCalmTheme.ColorSurface
            }

            Dim navItems As String() = {"Dashboard", "Document Registry", "SG Directives", "Search & Storage", "User & RFID Admin", "Audit Trail"}

            For i As Integer = 0 To navItems.Length - 1
                Dim idx As Integer = i
                Dim btn As New Button With {
                    .Text = "  " & navItems(i),
                    .Size = New Size(216, 44),
                    .Margin = New Padding(0, 0, 0, 6),
                    .FlatStyle = FlatStyle.Flat,
                    .Font = CivicCalmTheme.FontBody,
                    .TextAlign = ContentAlignment.MiddleLeft,
                    .Cursor = Cursors.Hand,
                    .ForeColor = CivicCalmTheme.ColorInkMuted,
                    .BackColor = CivicCalmTheme.ColorSurface
                }
                btn.FlatAppearance.BorderSize = 1
                btn.FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder
                btn.FlatAppearance.MouseOverBackColor = CivicCalmTheme.ColorWell

                AddHandler btn.Click, Sub() SwitchNavView(idx)
                navButtons.Add(btn)
                pnlNavStack.Controls.Add(btn)
            Next

            pnlSidebar.Controls.Add(pnlNavStack)
            pnlSidebar.Controls.Add(pnlBrand)
            Me.Controls.Add(pnlSidebar)
        End Sub

        Private Sub SwitchNavView(index As Integer)
            activeNavIndex = index
            For i As Integer = 0 To navButtons.Count - 1
                If i = index Then
                    navButtons(i).BackColor = CivicCalmTheme.ColorPrimarySoft
                    navButtons(i).ForeColor = CivicCalmTheme.ColorPrimary
                    navButtons(i).Font = CivicCalmTheme.FontFieldLabel
                    navButtons(i).FlatAppearance.BorderColor = CivicCalmTheme.ColorPrimary
                Else
                    navButtons(i).BackColor = CivicCalmTheme.ColorSurface
                    navButtons(i).ForeColor = CivicCalmTheme.ColorInkMuted
                    navButtons(i).Font = CivicCalmTheme.FontBody
                    navButtons(i).FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder
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

        Public Shared Sub ApplyGridStyle(dgv As DataGridView)
            DataGridStyler.ApplyCivicStyle(dgv)
        End Sub

        Private Sub SetupHeader()
            pnlHeader = New Panel With {
                .Dock = DockStyle.Top,
                .Height = 70,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(20, 12, 20, 12)
            }

            Dim pnlBorderBottom As New Panel With {
                .Dock = DockStyle.Bottom,
                .Height = 1,
                .BackColor = CivicCalmTheme.ColorBorder
            }
            pnlHeader.Controls.Add(pnlBorderBottom)

            Dim tblHeader As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 2,
                .RowCount = 1,
                .BackColor = CivicCalmTheme.ColorSurface
            }
            tblHeader.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 60.0F))
            tblHeader.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 40.0F))

            Dim pnlTitles As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.TopDown,
                .WrapContents = False,
                .BackColor = CivicCalmTheme.ColorSurface
            }

            lblTitle = New Label With {
                .Text = "OFFICE OF THE SECRETARY-GENERAL : DOCUMENT TRACKING DESK",
                .Font = CivicCalmTheme.FontFormTitle,
                .ForeColor = CivicCalmTheme.ColorInk,
                .AutoSize = True,
                .Margin = New Padding(0, 0, 0, 4)
            }

            lblUserBadge = New Label With {
                .Text = "[ RFID Logged Out: Access Restricted ]",
                .Font = CivicCalmTheme.FontMicrocopy,
                .ForeColor = CivicCalmTheme.ColorDanger,
                .AutoSize = True
            }
            pnlTitles.Controls.AddRange(New Control() {lblTitle, lblUserBadge})

            Dim pnlActions As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.RightToLeft,
                .WrapContents = False,
                .BackColor = CivicCalmTheme.ColorSurface
            }

            btnLogout = New Button With {
                .Text = "&Logout",
                .Size = New Size(95, 36),
                .BackColor = CivicCalmTheme.ColorWell,
                .ForeColor = CivicCalmTheme.ColorInk,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(8, 0, 0, 0)
            }
            btnLogout.FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder
            AddHandler btnLogout.Click, Sub() AuthenticateUser("")

            btnScanRFID = New Button With {
                .Text = "&Tap RFID Badge",
                .Size = New Size(160, 36),
                .BackColor = CivicCalmTheme.ColorPrimary,
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(8, 0, 0, 0)
            }
            btnScanRFID.FlatAppearance.BorderSize = 0
            AddHandler btnScanRFID.Click, Sub() ShowRFIDLoginDialog()

            pnlActions.Controls.AddRange(New Control() {btnLogout, btnScanRFID})

            tblHeader.Controls.Add(pnlTitles, 0, 0)
            tblHeader.Controls.Add(pnlActions, 1, 0)

            pnlHeader.Controls.Add(tblHeader)
            Me.Controls.Add(pnlHeader)
        End Sub

        Private Sub ShowRFIDLoginDialog()
            Using dlg As New FormLogin()
                If dlg.ShowDialog(Me) = DialogResult.OK AndAlso Not String.IsNullOrEmpty(dlg.ScannedUID) Then
                    AuthenticateUser(dlg.ScannedUID)
                End If
            End Using
        End Sub

        Public Sub AuthenticateUser(uid As String)
            If String.IsNullOrEmpty(uid) Then
                CurrentUser = Nothing
                CurrentSession = Nothing
                lblUserBadge.Text = "[ RFID Logged Out: Access Restricted ]"
                lblUserBadge.ForeColor = CivicCalmTheme.ColorDanger
                lblStatusMessage.Text = "User logged out."
            Else
                Dim user = EmbeddedDB.AuthenticateRFID(uid)
                If user IsNot Nothing Then
                    CurrentUser = user
                    Dim isGlobal As Boolean = (user("Role").ToString() = "Secretary-General" OrElse user("Role").ToString() = "System Administrator" OrElse user("Role").ToString() = "OSG Chief")
                    lblUserBadge.Text = String.Format("AUTHENTICATED: {0} [{1}] - {2}", user("FullName").ToString().ToUpperInvariant(), user("Role").ToString().ToUpperInvariant(), If(isGlobal, "GLOBAL ACCESS", "STAFF VIEW"))
                    lblUserBadge.ForeColor = CivicCalmTheme.ColorPrimary
                    lblStatusMessage.Text = "Authenticated: " & user("FullName").ToString()
                    EmbeddedDB.LogAudit(user("FullName").ToString(), "RFID Badge Tap Authenticated [UID: " & uid & "]")
                Else
                    lblStatusMessage.Text = "Access Denied: Unrecognized RFID card [" & uid & "]"
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
            viewDashboard = New Panel With {.Dock = DockStyle.Fill, .BackColor = CivicCalmTheme.ColorCanvas}

            Dim pnlCards As New TableLayoutPanel With {
                .Dock = DockStyle.Top,
                .Height = 100,
                .ColumnCount = 4,
                .RowCount = 1,
                .Padding = New Padding(0, 0, 0, 16)
            }
            pnlCards.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 25.0F))
            pnlCards.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 25.0F))
            pnlCards.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 25.0F))
            pnlCards.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 25.0F))

            lblStatTotalDocs = CreateStatCard(pnlCards, 0, "TOTAL DOCUMENTS", "0", CivicCalmTheme.ColorPrimary)
            lblStatDirectives = CreateStatCard(pnlCards, 1, "SG DIRECTIVES", "0", CivicCalmTheme.ColorAccentSG)
            lblStatActiveRoute = CreateStatCard(pnlCards, 2, "ACTIVE ROUTINGS", "0", ColorTranslator.FromHtml("#0284C7"))
            lblStatVaultStorage = CreateStatCard(pnlCards, 3, "PHYSICAL VAULT ITEMS", "0", ColorTranslator.FromHtml("#D97706"))

            Dim pnlGridCard As New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(12)
            }

            Dim lblRecHeader As New Label With {
                .Text = "RECENT PARLIAMENTARY DOCUMENTS",
                .Font = CivicCalmTheme.FontSectionHeader,
                .ForeColor = CivicCalmTheme.ColorInk,
                .Dock = DockStyle.Top,
                .Height = 32
            }

            dgvDashRecent = New DataGridView With {
                .Dock = DockStyle.Fill,
                .ReadOnly = True,
                .AllowUserToAddRows = False,
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            }
            DataGridStyler.ApplyCivicStyle(dgvDashRecent)
            AddHandler dgvDashRecent.CellDoubleClick, Sub(s, e) If e.RowIndex >= 0 Then OpenSelectedDocumentDetail(dgvDashRecent)

            lblDashWatermark = CreateGridWatermark(pnlGridCard)

            pnlGridCard.Controls.Add(dgvDashRecent)
            pnlGridCard.Controls.Add(lblRecHeader)

            viewDashboard.Controls.Add(pnlGridCard)
            viewDashboard.Controls.Add(pnlCards)
        End Sub

        Private Function CreateStatCard(parent As TableLayoutPanel, colIndex As Integer, title As String, initVal As String, accentBg As Color) As Label
            Dim pnlCard As New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Margin = New Padding(4)
            }
            Dim pnlBar As New Panel With {
                .Dock = DockStyle.Top,
                .Height = 4,
                .BackColor = accentBg
            }
            Dim lblT As New Label With {
                .Text = title,
                .Font = CivicCalmTheme.FontMicrocopy,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .Dock = DockStyle.Top,
                .Height = 24,
                .Padding = New Padding(12, 6, 0, 0)
            }
            Dim lblV As New Label With {
                .Text = initVal,
                .Font = New Font("Segoe UI", 16.0F, FontStyle.Bold),
                .ForeColor = CivicCalmTheme.ColorInk,
                .Dock = DockStyle.Fill,
                .Padding = New Padding(12, 0, 0, 0),
                .TextAlign = ContentAlignment.MiddleLeft
            }

            pnlCard.Controls.AddRange(New Control() {lblV, lblT, pnlBar})
            parent.Controls.Add(pnlCard, colIndex, 0)
            Return lblV
        End Function

        Private Sub SetupRegistryView()
            viewRegistry = New Panel With {.Dock = DockStyle.Fill, .BackColor = CivicCalmTheme.ColorCanvas}

            Dim tblMain As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 2,
                .RowCount = 1
            }
            tblMain.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 400.0F))
            tblMain.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))

            ' Left: Form Card Panel
            Dim pnlFormCard As New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(16),
                .Margin = New Padding(0, 0, 16, 0)
            }

            Dim pnlFormScroll As New Panel With {
                .Dock = DockStyle.Fill,
                .AutoScroll = True,
                .BackColor = CivicCalmTheme.ColorSurface
            }

            txtTitle = New TextBox With {.Height = 28, .Text = "Draft Resolution on BTA Regional Governance", .BackColor = CivicCalmTheme.ColorSurface, .ForeColor = CivicCalmTheme.ColorInk, .BorderStyle = BorderStyle.FixedSingle}
            cmbDocType = New ComboBox With {.DropDownStyle = ComboBoxStyle.DropDownList, .Height = 28, .BackColor = CivicCalmTheme.ColorSurface, .ForeColor = CivicCalmTheme.ColorInk}
            cmbDocType.Items.AddRange(New Object() {"Resolution", "Parliament Bill", "Committee Report", "Executive Communication", "Memorandum", "Endorsement", "Journal Entry"})
            cmbDocType.SelectedIndex = 0

            txtOrigin = New TextBox With {.Height = 28, .Text = "Office of MP Yasser", .BackColor = CivicCalmTheme.ColorSurface, .ForeColor = CivicCalmTheme.ColorInk, .BorderStyle = BorderStyle.FixedSingle}
            txtDest = New TextBox With {.Height = 28, .Text = "Office of the Secretary-General", .BackColor = CivicCalmTheme.ColorSurface, .ForeColor = CivicCalmTheme.ColorInk, .BorderStyle = BorderStyle.FixedSingle}
            txtCabinet = New TextBox With {.Height = 28, .Text = "CAB-A", .BackColor = CivicCalmTheme.ColorSurface, .ForeColor = CivicCalmTheme.ColorInk, .BorderStyle = BorderStyle.FixedSingle}
            txtShelf = New TextBox With {.Height = 28, .Text = "S-2", .BackColor = CivicCalmTheme.ColorSurface, .ForeColor = CivicCalmTheme.ColorInk, .BorderStyle = BorderStyle.FixedSingle}
            txtBox = New TextBox With {.Height = 28, .Text = "BOX-03", .BackColor = CivicCalmTheme.ColorSurface, .ForeColor = CivicCalmTheme.ColorInk, .BorderStyle = BorderStyle.FixedSingle}
            txtGDrive = New TextBox With {.Height = 28, .Text = "https://drive.google.com/file/d/bta-doc-2026/view", .BackColor = CivicCalmTheme.ColorSurface, .ForeColor = CivicCalmTheme.ColorInk, .BorderStyle = BorderStyle.FixedSingle}
            cmbAssignedStaff = New ComboBox With {.DropDownStyle = ComboBoxStyle.DropDownList, .Height = 28, .BackColor = CivicCalmTheme.ColorSurface, .ForeColor = CivicCalmTheme.ColorInk}

            Dim pnlFormFlow As New FlowLayoutPanel With {
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .FlowDirection = FlowDirection.TopDown,
                .WrapContents = False,
                .BackColor = CivicCalmTheme.ColorSurface
            }

            Dim AddField = Sub(labelText As String, ctrl As Control)
                               Dim lbl As New Label With {
                                   .Text = labelText,
                                   .AutoSize = True,
                                   .Font = CivicCalmTheme.FontFieldLabel,
                                   .ForeColor = CivicCalmTheme.ColorInkMuted,
                                   .Margin = New Padding(0, 8, 0, 2)
                               }
                               ctrl.Width = 340
                               ctrl.Margin = New Padding(0, 0, 0, 6)
                               pnlFormFlow.Controls.Add(lbl)
                               pnlFormFlow.Controls.Add(ctrl)
                           End Sub

            AddField("Document Title:", txtTitle)
            AddField("Document Type:", cmbDocType)
            AddField("Originating Office:", txtOrigin)
            AddField("Destination Office:", txtDest)
            AddField("Cabinet Landmark ID:", txtCabinet)
            AddField("Shelf Landmark No:", txtShelf)
            AddField("Box Landmark Code:", txtBox)
            AddField("Google Drive Soft Copy URL:", txtGDrive)
            AddField("Assigned OSG Staff:", cmbAssignedStaff)

            btnRegister = New Button With {
                .Text = "&Register Parliamentary Document",
                .Size = New Size(340, 42),
                .BackColor = CivicCalmTheme.ColorPrimary,
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(0, 12, 0, 12)
            }
            btnRegister.FlatAppearance.BorderSize = 0
            AddHandler btnRegister.Click, AddressOf OnRegisterDocument
            pnlFormFlow.Controls.Add(btnRegister)

            pnlFormScroll.Controls.Add(pnlFormFlow)
            pnlFormCard.Controls.Add(pnlFormScroll)

            ' Right: Grid Card Panel
            Dim pnlGridCard As New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(12)
            }

            Dim pnlToolbar As New Panel With {
                .Dock = DockStyle.Top,
                .Height = 44,
                .BackColor = CivicCalmTheme.ColorSurface
            }
            btnViewRegistryDetail = New Button With {
                .Text = "&View Details & History",
                .Size = New Size(220, 34),
                .BackColor = CivicCalmTheme.ColorWell,
                .ForeColor = CivicCalmTheme.ColorInk,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand
            }
            btnViewRegistryDetail.FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder
            AddHandler btnViewRegistryDetail.Click, Sub() OpenSelectedDocumentDetail(dgvRegistry)
            pnlToolbar.Controls.Add(btnViewRegistryDetail)

            dgvRegistry = New DataGridView With {
                .Dock = DockStyle.Fill,
                .ReadOnly = True,
                .AllowUserToAddRows = False,
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            }
            DataGridStyler.ApplyCivicStyle(dgvRegistry)
            AddHandler dgvRegistry.CellDoubleClick, Sub(s, e) If e.RowIndex >= 0 Then OpenSelectedDocumentDetail(dgvRegistry)

            lblRegistryWatermark = CreateGridWatermark(pnlGridCard)

            pnlGridCard.Controls.Add(dgvRegistry)
            pnlGridCard.Controls.Add(pnlToolbar)

            tblMain.Controls.Add(pnlFormCard, 0, 0)
            tblMain.Controls.Add(pnlGridCard, 1, 0)
            viewRegistry.Controls.Add(tblMain)
        End Sub

        Private Sub OnRegisterDocument(sender As Object, e As EventArgs)
            epValidation.Clear()

            If CurrentUser Is Nothing Then
                lblStatusMessage.Text = "Authentication Required: Tap RFID card to register documents."
                MessageBox.Show("RFID Authentication Required to Register Documents.", "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            If String.IsNullOrWhiteSpace(txtTitle.Text) Then
                epValidation.SetError(txtTitle, "Please enter Document Title.")
                lblStatusMessage.Text = "Validation Error: Document Title is required."
                txtTitle.Focus()
                Return
            End If

            Dim errUrlMsg As String = ""
            If Not EmbeddedDB.ValidateGDriveURL(txtGDrive.Text, errUrlMsg) Then
                epValidation.SetError(txtGDrive, errUrlMsg)
                lblStatusMessage.Text = "Invalid Soft Copy URL: " & errUrlMsg
                txtGDrive.Focus()
                Return
            End If

            Dim docType = cmbDocType.SelectedItem.ToString()
            Dim code = EmbeddedDB.GenerateDocCode(docType)
            Dim assigned = If(cmbAssignedStaff.SelectedItem IsNot Nothing, cmbAssignedStaff.SelectedItem.ToString(), "")

            EmbeddedDB.AddDocument(code, docType, txtTitle.Text.Trim(), txtOrigin.Text.Trim(), txtDest.Text.Trim(), txtCabinet.Text.Trim(), txtShelf.Text.Trim(), txtBox.Text.Trim(), txtGDrive.Text.Trim(), "LOGGED", assigned)
            EmbeddedDB.LogAudit(CurrentUser("FullName").ToString(), String.Format("Registered New Document [{0}] - {1}", code, txtTitle.Text.Trim()))

            lblStatusMessage.Text = String.Format("Document Registered: {0} [{1}]", code, txtTitle.Text.Trim())
            txtTitle.Text = ""
            RefreshActiveTabGrid()
        End Sub

        Private Sub SetupDirectivesView()
            viewDirectives = New Panel With {.Dock = DockStyle.Fill, .BackColor = CivicCalmTheme.ColorCanvas}

            Dim pnlTopCard As New Panel With {
                .Dock = DockStyle.Top,
                .Height = 120,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(16),
                .Margin = New Padding(0, 0, 0, 16)
            }

            Dim tblDir As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 4,
                .RowCount = 2
            }
            tblDir.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 35.0F))
            tblDir.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 30.0F))
            tblDir.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 20.0F))
            tblDir.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 15.0F))

            cmbDirDocs = New ComboBox With {.Dock = DockStyle.Fill, .DropDownStyle = ComboBoxStyle.DropDownList, .BackColor = CivicCalmTheme.ColorSurface, .ForeColor = CivicCalmTheme.ColorInk}
            cmbDirective = New ComboBox With {.Dock = DockStyle.Fill, .DropDownStyle = ComboBoxStyle.DropDownList, .BackColor = CivicCalmTheme.ColorSurface, .ForeColor = CivicCalmTheme.ColorInk}
            cmbDirective.Items.AddRange(New Object() {"For Immediate Action", "Referred to Committee on Rules", "Forwarded for Speaker Signature", "Under OSG Administrative Review", "Approved & Archived"})
            cmbDirective.SelectedIndex = 0

            cmbDirAssign = New ComboBox With {.Dock = DockStyle.Fill, .DropDownStyle = ComboBoxStyle.DropDownList, .BackColor = CivicCalmTheme.ColorSurface, .ForeColor = CivicCalmTheme.ColorInk}
            txtDirNotes = New TextBox With {.Dock = DockStyle.Fill, .Text = "Priority routing per Secretary-General directive.", .BackColor = CivicCalmTheme.ColorSurface, .ForeColor = CivicCalmTheme.ColorInk, .BorderStyle = BorderStyle.FixedSingle}

            btnApplyDirective = New Button With {
                .Text = "&Log Action Directive",
                .Dock = DockStyle.Fill,
                .BackColor = CivicCalmTheme.ColorPrimary,
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand
            }
            btnApplyDirective.FlatAppearance.BorderSize = 0
            AddHandler btnApplyDirective.Click, AddressOf OnApplyDirective

            tblDir.Controls.Add(New Label With {.Text = "Select Document:", .Font = CivicCalmTheme.FontFieldLabel, .ForeColor = CivicCalmTheme.ColorInkMuted, .Dock = DockStyle.Fill}, 0, 0)
            tblDir.Controls.Add(New Label With {.Text = "SG Directive:", .Font = CivicCalmTheme.FontFieldLabel, .ForeColor = CivicCalmTheme.ColorInkMuted, .Dock = DockStyle.Fill}, 1, 0)
            tblDir.Controls.Add(New Label With {.Text = "Reassign Staff:", .Font = CivicCalmTheme.FontFieldLabel, .ForeColor = CivicCalmTheme.ColorInkMuted, .Dock = DockStyle.Fill}, 2, 0)
            tblDir.Controls.Add(New Label With {.Text = "", .Dock = DockStyle.Fill}, 3, 0)

            tblDir.Controls.Add(cmbDirDocs, 0, 1)
            tblDir.Controls.Add(cmbDirective, 1, 1)
            tblDir.Controls.Add(cmbDirAssign, 2, 1)
            tblDir.Controls.Add(btnApplyDirective, 3, 1)

            pnlTopCard.Controls.Add(tblDir)

            Dim pnlGridCard As New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(12),
                .Margin = New Padding(0, 16, 0, 0)
            }

            dgvDirectives = New DataGridView With {
                .Dock = DockStyle.Fill,
                .ReadOnly = True,
                .AllowUserToAddRows = False,
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            }
            DataGridStyler.ApplyCivicStyle(dgvDirectives)

            lblDirectivesWatermark = CreateGridWatermark(pnlGridCard)

            pnlGridCard.Controls.Add(dgvDirectives)

            viewDirectives.Controls.Add(pnlGridCard)
            viewDirectives.Controls.Add(pnlTopCard)
        End Sub

        Private Sub OnApplyDirective(sender As Object, e As EventArgs)
            If CurrentUser Is Nothing Then
                lblStatusMessage.Text = "Authentication Required to Issue Directives."
                MessageBox.Show("RFID Authentication Required.", "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            If cmbDirDocs.SelectedItem Is Nothing Then
                lblStatusMessage.Text = "Validation Error: Please select a document."
                Return
            End If

            Dim docStr = cmbDirDocs.SelectedItem.ToString()
            Dim docId As Integer = CInt(docStr.Split(":"c)(0).Replace("ID ", "").Trim())
            Dim directive = cmbDirective.SelectedItem.ToString()
            Dim assign = If(cmbDirAssign.SelectedItem IsNot Nothing, cmbDirAssign.SelectedItem.ToString(), "")

            EmbeddedDB.AddDirective(docId, directive, assign, txtDirNotes.Text.Trim(), CurrentUser("FullName").ToString())
            EmbeddedDB.LogAudit(CurrentUser("FullName").ToString(), String.Format("Applied SG Directive [{0}] to Doc ID #{1}", directive, docId))

            lblStatusMessage.Text = String.Format("Action Directive Logged: [{0}] applied to Doc ID #{1}", directive, docId)
            RefreshActiveTabGrid()
        End Sub

        Private Sub SetupSearchView()
            viewSearch = New Panel With {.Dock = DockStyle.Fill, .BackColor = CivicCalmTheme.ColorCanvas}

            Dim pnlSearchCard As New Panel With {
                .Dock = DockStyle.Top,
                .Height = 65,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(12),
                .Margin = New Padding(0, 0, 0, 16)
            }

            Dim flwSearch As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.LeftToRight,
                .WrapContents = False,
                .BackColor = CivicCalmTheme.ColorSurface
            }

            Dim lblKey As New Label With {
                .Text = "Search Keyword:",
                .Font = CivicCalmTheme.FontFieldLabel,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .AutoSize = True,
                .Margin = New Padding(0, 8, 8, 0)
            }

            txtSearchKey = New TextBox With {
                .Width = 320,
                .Height = 32,
                .BackColor = CivicCalmTheme.ColorSurface,
                .ForeColor = CivicCalmTheme.ColorInk,
                .BorderStyle = BorderStyle.FixedSingle,
                .Margin = New Padding(0, 4, 8, 0)
            }

            btnSearch = New Button With {
                .Text = "&Search",
                .Size = New Size(95, 32),
                .BackColor = CivicCalmTheme.ColorPrimary,
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(0, 4, 8, 0)
            }
            btnSearch.FlatAppearance.BorderSize = 0

            btnClearSearch = New Button With {
                .Text = "&Clear Filters",
                .Size = New Size(110, 32),
                .BackColor = CivicCalmTheme.ColorWell,
                .ForeColor = CivicCalmTheme.ColorInk,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontBody,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(0, 4, 8, 0)
            }
            btnClearSearch.FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder

            btnViewSearchDetail = New Button With {
                .Text = "&View Details",
                .Size = New Size(130, 32),
                .BackColor = CivicCalmTheme.ColorWell,
                .ForeColor = CivicCalmTheme.ColorInk,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(0, 4, 8, 0)
            }
            btnViewSearchDetail.FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder

            btnOpenPDF = New Button With {
                .Text = "&Launch PDF",
                .Size = New Size(120, 32),
                .BackColor = CivicCalmTheme.ColorWell,
                .ForeColor = CivicCalmTheme.ColorInk,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(0, 4, 0, 0)
            }
            btnOpenPDF.FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder

            AddHandler btnSearch.Click, AddressOf OnSearch
            AddHandler btnClearSearch.Click, Sub()
                                                 txtSearchKey.Text = ""
                                                 OnSearch(Nothing, EventArgs.Empty)
                                             End Sub
            AddHandler btnViewSearchDetail.Click, Sub() OpenSelectedDocumentDetail(dgvSearch)
            AddHandler btnOpenPDF.Click, AddressOf OnOpenPDF

            flwSearch.Controls.AddRange(New Control() {lblKey, txtSearchKey, btnSearch, btnClearSearch, btnViewSearchDetail, btnOpenPDF})
            pnlSearchCard.Controls.Add(flwSearch)

            Dim pnlGridCard As New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(12),
                .Margin = New Padding(0, 16, 0, 0)
            }

            dgvSearch = New DataGridView With {
                .Dock = DockStyle.Fill,
                .ReadOnly = True,
                .AllowUserToAddRows = False,
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            }
            DataGridStyler.ApplyCivicStyle(dgvSearch)
            AddHandler dgvSearch.CellDoubleClick, Sub(s, e) If e.RowIndex >= 0 Then OpenSelectedDocumentDetail(dgvSearch)

            lblSearchWatermark = CreateGridWatermark(pnlGridCard)

            pnlGridCard.Controls.Add(dgvSearch)

            viewSearch.Controls.Add(pnlGridCard)
            viewSearch.Controls.Add(pnlSearchCard)
        End Sub

        Private Sub OnSearch(sender As Object, e As EventArgs)
            Dim visibleTable = EmbeddedDB.GetVisibleDocuments(CurrentUser)
            Dim q = txtSearchKey.Text.Trim().Replace("'", "''")
            If String.IsNullOrEmpty(q) Then
                dgvSearch.DataSource = visibleTable
                lblStatusMessage.Text = "Displaying all visible documents."
            Else
                Dim filter As String = String.Format("DocCode LIKE '%{0}%' OR Title LIKE '%{0}%' OR DocType LIKE '%{0}%' OR CabinetID LIKE '%{0}%' OR OriginatingOffice LIKE '%{0}%' OR DestinationOffice LIKE '%{0}%' OR AssignedStaff LIKE '%{0}%'", q)
                Dim dv As New DataView(visibleTable)
                dv.RowFilter = filter
                dgvSearch.DataSource = dv.ToTable()
                lblStatusMessage.Text = String.Format("Search complete for '{0}'.", txtSearchKey.Text.Trim())
            End If

            If dgvSearch.Rows.Count = 0 Then
                DataGridStyler.SetEmptyState(dgvSearch, lblSearchWatermark, "No documents match the current filter criteria. Press Alt+C to clear filters.")
            Else
                DataGridStyler.SetPopulatedState(dgvSearch, lblSearchWatermark)
            End If
            lblStatusCount.Text = dgvSearch.Rows.Count.ToString() & " Records"
        End Sub

        Private Sub OnOpenPDF(sender As Object, e As EventArgs)
            If dgvSearch.CurrentRow Is Nothing Then
                lblStatusMessage.Text = "Please select a document row from the grid first."
                Return
            End If

            Dim url As String = dgvSearch.CurrentRow.Cells("GDriveURL").Value.ToString()
            If Not String.IsNullOrEmpty(url) Then
                Dim errUrl As String = ""
                If Not EmbeddedDB.ValidateGDriveURL(url, errUrl) Then
                    lblStatusMessage.Text = "Security Error: " & errUrl
                    Return
                End If
                Try
                    Process.Start(New ProcessStartInfo With {.FileName = url, .UseShellExecute = True})
                    If CurrentUser IsNot Nothing Then
                        EmbeddedDB.LogAudit(CurrentUser("FullName").ToString(), "Launched Google Drive Soft Copy PDF: " & url)
                    End If
                    lblStatusMessage.Text = "Launched PDF document in default viewer."
                Catch ex As Exception
                    lblStatusMessage.Text = "Unable to launch PDF viewer: " & ex.Message
                End Try
            Else
                lblStatusMessage.Text = "No Google Drive PDF URL associated with this document."
            End If
        End Sub

        Private Sub SetupAdminView()
            viewAdmin = New Panel With {.Dock = DockStyle.Fill, .BackColor = CivicCalmTheme.ColorCanvas}

            Dim tblMain As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 2,
                .RowCount = 1
            }
            tblMain.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 380.0F))
            tblMain.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))

            Dim pnlFormCard As New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(16),
                .Margin = New Padding(0, 0, 16, 0)
            }

            Dim pnlFormFlow As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.TopDown,
                .WrapContents = False,
                .BackColor = CivicCalmTheme.ColorSurface
            }

            Dim lblHeader As New Label With {
                .Text = "REGISTER USER & RFID BADGE",
                .Font = CivicCalmTheme.FontSectionHeader,
                .ForeColor = CivicCalmTheme.ColorInk,
                .AutoSize = True,
                .Margin = New Padding(0, 0, 0, 12)
            }

            txtNewUserName = New TextBox With {.Width = 320, .Height = 28, .Text = "New Staff Member", .BackColor = CivicCalmTheme.ColorSurface, .ForeColor = CivicCalmTheme.ColorInk, .BorderStyle = BorderStyle.FixedSingle}
            cmbNewUserRole = New ComboBox With {.Width = 320, .Height = 28, .DropDownStyle = ComboBoxStyle.DropDownList, .BackColor = CivicCalmTheme.ColorSurface, .ForeColor = CivicCalmTheme.ColorInk}
            cmbNewUserRole.Items.AddRange(New Object() {"Secretary-General", "OSG Chief", "System Administrator", "Administrative Staff"})
            cmbNewUserRole.SelectedIndex = 3

            txtNewUserUID = New TextBox With {.Width = 320, .Height = 28, .Text = "44D4E555", .BackColor = CivicCalmTheme.ColorSurface, .ForeColor = CivicCalmTheme.ColorInk, .BorderStyle = BorderStyle.FixedSingle}

            btnAddUser = New Button With {
                .Text = "&Save User & RFID Smart Card",
                .Size = New Size(320, 42),
                .BackColor = CivicCalmTheme.ColorPrimary,
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(0, 16, 0, 0)
            }
            btnAddUser.FlatAppearance.BorderSize = 0
            AddHandler btnAddUser.Click, AddressOf OnAddUser

            pnlFormFlow.Controls.Add(lblHeader)
            pnlFormFlow.Controls.Add(New Label With {.Text = "Full Name:", .Font = CivicCalmTheme.FontFieldLabel, .ForeColor = CivicCalmTheme.ColorInkMuted, .Margin = New Padding(0, 8, 0, 2)})
            pnlFormFlow.Controls.Add(txtNewUserName)
            pnlFormFlow.Controls.Add(New Label With {.Text = "System Role:", .Font = CivicCalmTheme.FontFieldLabel, .ForeColor = CivicCalmTheme.ColorInkMuted, .Margin = New Padding(0, 8, 0, 2)})
            pnlFormFlow.Controls.Add(cmbNewUserRole)
            pnlFormFlow.Controls.Add(New Label With {.Text = "RFID Smart Card UID (Hex):", .Font = CivicCalmTheme.FontFieldLabel, .ForeColor = CivicCalmTheme.ColorInkMuted, .Margin = New Padding(0, 8, 0, 2)})
            pnlFormFlow.Controls.Add(txtNewUserUID)
            pnlFormFlow.Controls.Add(btnAddUser)

            pnlFormCard.Controls.Add(pnlFormFlow)

            Dim pnlGridCard As New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(12)
            }

            dgvUsers = New DataGridView With {
                .Dock = DockStyle.Fill,
                .ReadOnly = True,
                .AllowUserToAddRows = False,
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            }
            DataGridStyler.ApplyCivicStyle(dgvUsers)

            lblUsersWatermark = CreateGridWatermark(pnlGridCard)

            pnlGridCard.Controls.Add(dgvUsers)

            tblMain.Controls.Add(pnlFormCard, 0, 0)
            tblMain.Controls.Add(pnlGridCard, 1, 0)
            viewAdmin.Controls.Add(tblMain)
        End Sub

        Private Sub OnAddUser(sender As Object, e As EventArgs)
            If CurrentUser Is Nothing OrElse CurrentUser("Role").ToString() <> "System Administrator" Then
                lblStatusMessage.Text = "Access Denied: System Administrator privileges required."
                MessageBox.Show("System Administrator Privileges Required.", "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            If String.IsNullOrWhiteSpace(txtNewUserName.Text) OrElse String.IsNullOrWhiteSpace(txtNewUserUID.Text) Then
                lblStatusMessage.Text = "Validation Error: Please enter Full Name and RFID Card UID."
                Return
            End If

            EmbeddedDB.AddUser(txtNewUserUID.Text.Trim(), txtNewUserName.Text.Trim(), cmbNewUserRole.SelectedItem.ToString())
            EmbeddedDB.LogAudit(CurrentUser("FullName").ToString(), String.Format("Registered/Updated User [{0}] Role: {1} RFID: {2}", txtNewUserName.Text.Trim(), cmbNewUserRole.SelectedItem, txtNewUserUID.Text.Trim()))

            lblStatusMessage.Text = String.Format("User Registered: {0} [{1}]", txtNewUserName.Text.Trim(), cmbNewUserRole.SelectedItem.ToString())
            PopulateStaffDropdowns()
            RefreshActiveTabGrid()
        End Sub

        Private Sub SetupAuditView()
            viewAudit = New Panel With {.Dock = DockStyle.Fill, .BackColor = CivicCalmTheme.ColorCanvas}

            Dim pnlAuditCard As New Panel With {
                .Dock = DockStyle.Top,
                .Height = 65,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(12),
                .Margin = New Padding(0, 0, 0, 16)
            }

            Dim flwAudit As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.LeftToRight,
                .WrapContents = False,
                .BackColor = CivicCalmTheme.ColorSurface
            }

            Dim lblF As New Label With {
                .Text = "Filter Audit Trail:",
                .Font = CivicCalmTheme.FontFieldLabel,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .AutoSize = True,
                .Margin = New Padding(0, 8, 8, 0)
            }

            txtAuditSearch = New TextBox With {
                .Width = 320,
                .Height = 32,
                .BackColor = CivicCalmTheme.ColorSurface,
                .ForeColor = CivicCalmTheme.ColorInk,
                .BorderStyle = BorderStyle.FixedSingle,
                .Margin = New Padding(0, 4, 8, 0)
            }

            btnAuditFilter = New Button With {
                .Text = "&Filter",
                .Size = New Size(95, 32),
                .BackColor = CivicCalmTheme.ColorPrimary,
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(0, 4, 8, 0)
            }
            btnAuditFilter.FlatAppearance.BorderSize = 0

            btnAuditReset = New Button With {
                .Text = "&Reset",
                .Size = New Size(95, 32),
                .BackColor = CivicCalmTheme.ColorWell,
                .ForeColor = CivicCalmTheme.ColorInk,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontBody,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(0, 4, 0, 0)
            }
            btnAuditReset.FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder

            AddHandler btnAuditFilter.Click, AddressOf OnFilterAudit
            AddHandler btnAuditReset.Click, Sub()
                                                txtAuditSearch.Text = ""
                                                OnFilterAudit(Nothing, EventArgs.Empty)
                                            End Sub

            flwAudit.Controls.AddRange(New Control() {lblF, txtAuditSearch, btnAuditFilter, btnAuditReset})
            pnlAuditCard.Controls.Add(flwAudit)

            Dim pnlGridCard As New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(12),
                .Margin = New Padding(0, 16, 0, 0)
            }

            dgvAudit = New DataGridView With {
                .Dock = DockStyle.Fill,
                .ReadOnly = True,
                .AllowUserToAddRows = False,
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            }
            DataGridStyler.ApplyCivicStyle(dgvAudit)

            lblAuditWatermark = CreateGridWatermark(pnlGridCard)

            pnlGridCard.Controls.Add(dgvAudit)

            viewAudit.Controls.Add(pnlGridCard)
            viewAudit.Controls.Add(pnlAuditCard)
        End Sub

        Private Sub OnFilterAudit(sender As Object, e As EventArgs)
            Dim q = txtAuditSearch.Text.Trim().Replace("'", "''")
            Dim dt = EmbeddedDB.DataSet.Tables("AuditTrail")
            If String.IsNullOrEmpty(q) Then
                dgvAudit.DataSource = dt
                lblStatusMessage.Text = "Displaying complete audit trail."
            Else
                Dim dv As New DataView(dt)
                dv.RowFilter = String.Format("UserName LIKE '%{0}%' OR ActionDescription LIKE '%{0}%'", q)
                dgvAudit.DataSource = dv.ToTable()
                lblStatusMessage.Text = String.Format("Audit log filtered for '{0}'.", txtAuditSearch.Text.Trim())
            End If

            If dgvAudit.Rows.Count = 0 Then
                DataGridStyler.SetEmptyState(dgvAudit, lblAuditWatermark, "No audit trail records match the search filter.")
            Else
                DataGridStyler.SetPopulatedState(dgvAudit, lblAuditWatermark)
            End If
            lblStatusCount.Text = dgvAudit.Rows.Count.ToString() & " Records"
        End Sub

        Private Function CreateGridWatermark(parent As Control) As Label
            Dim lbl As New Label With {
                .Text = "No documents match the current filter criteria. Press Alt+C to clear filters.",
                .Font = CivicCalmTheme.FontBody,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .TextAlign = ContentAlignment.MiddleCenter,
                .Dock = DockStyle.Fill,
                .Visible = False,
                .BackColor = CivicCalmTheme.ColorSurface
            }
            parent.Controls.Add(lbl)
            Return lbl
        End Function

        Public Sub OpenSelectedDocumentDetail(dgv As DataGridView)
            If dgv.CurrentRow Is Nothing Then
                lblStatusMessage.Text = "Please select a document row from the grid first."
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
            If dgvDashRecent.Rows.Count = 0 Then
                DataGridStyler.SetEmptyState(dgvDashRecent, lblDashWatermark)
            Else
                DataGridStyler.SetPopulatedState(dgvDashRecent, lblDashWatermark)
            End If

            Select Case activeNavIndex
                Case 1 ' Registry View
                    dgvRegistry.DataSource = visibleDocs
                    If dgvRegistry.Rows.Count = 0 Then
                        DataGridStyler.SetEmptyState(dgvRegistry, lblRegistryWatermark)
                    Else
                        DataGridStyler.SetPopulatedState(dgvRegistry, lblRegistryWatermark)
                    End If
                    lblStatusCount.Text = dgvRegistry.Rows.Count.ToString() & " Documents"
                Case 2 ' Directives View
                    cmbDirDocs.Items.Clear()
                    For Each row As DataRow In visibleDocs.Rows
                        cmbDirDocs.Items.Add(String.Format("ID {0}: [{1}] {2}", row("DocumentID"), row("DocCode"), row("Title")))
                    Next
                    If cmbDirDocs.Items.Count > 0 Then cmbDirDocs.SelectedIndex = 0
                    Dim dtDirectives = EmbeddedDB.DataSet.Tables("Directives")
                    dgvDirectives.DataSource = dtDirectives
                    If dgvDirectives.Rows.Count = 0 Then
                        DataGridStyler.SetEmptyState(dgvDirectives, lblDirectivesWatermark, "No active action directives found.")
                    Else
                        DataGridStyler.SetPopulatedState(dgvDirectives, lblDirectivesWatermark)
                    End If
                    lblStatusCount.Text = dgvDirectives.Rows.Count.ToString() & " Directives"
                Case 3 ' Search View
                    dgvSearch.DataSource = visibleDocs
                    If dgvSearch.Rows.Count = 0 Then
                        DataGridStyler.SetEmptyState(dgvSearch, lblSearchWatermark)
                    Else
                        DataGridStyler.SetPopulatedState(dgvSearch, lblSearchWatermark)
                    End If
                    lblStatusCount.Text = dgvSearch.Rows.Count.ToString() & " Documents"
                Case 4 ' Admin View
                    Dim dtUsers = EmbeddedDB.DataSet.Tables("Users")
                    dgvUsers.DataSource = dtUsers
                    If dgvUsers.Rows.Count = 0 Then
                        DataGridStyler.SetEmptyState(dgvUsers, lblUsersWatermark, "No users registered.")
                    Else
                        DataGridStyler.SetPopulatedState(dgvUsers, lblUsersWatermark)
                    End If
                    lblStatusCount.Text = dgvUsers.Rows.Count.ToString() & " Users"
                Case 5 ' Audit View
                    Dim dtAudit = EmbeddedDB.DataSet.Tables("AuditTrail")
                    dgvAudit.DataSource = dtAudit
                    If dgvAudit.Rows.Count = 0 Then
                        DataGridStyler.SetEmptyState(dgvAudit, lblAuditWatermark, "No audit trail events recorded.")
                    Else
                        DataGridStyler.SetPopulatedState(dgvAudit, lblAuditWatermark)
                    End If
                    lblStatusCount.Text = dgvAudit.Rows.Count.ToString() & " Audit Events"
            End Select

            lblStatusClock.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        End Sub
    End Class
End Namespace
