Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Drawing
Imports System.Windows.Forms

Namespace BTA_OSG
    Partial Public Class FormMain
        Inherits Form

        Public CurrentUser As DataRow = Nothing

        ' Compose the whole control tree offscreen on repaint (see UiBuffering.WsExComposited).
        ' Without this, minimize/restore lets each child HWND paint separately and stale copies
        ' of grids and panels linger at their pre-restore positions.
        Protected Overrides ReadOnly Property CreateParams As CreateParams
            Get
                Dim cp As CreateParams = MyBase.CreateParams
                cp.ExStyle = cp.ExStyle Or UiBuffering.WsExComposited
                Return cp
            End Get
        End Property

        ' Layout Panels
        Private pnlSidebar As Panel
        Private pnlHeader As Panel
        Private pnlContent As Panel
        Private stsFooter As StatusStrip
        Private lblStatusMessage As ToolStripStatusLabel
        Private lblStatusCount As ToolStripStatusLabel
        Private lblStatusClock As ToolStripStatusLabel
        Private epValidation As ErrorProvider
        Private tmrClock As Timer
        Private tmrPortalPoll As Timer
        Private tmrSqlSync As Timer

        ' One SQL snapshot pull at a time: a slow host skips ticks instead of queueing them.
        Private _sqlSyncInFlight As Integer

        ' Header Controls
        Private lblTitle As Label
        Private lblUserBadge As Label
        Private btnScanRFID As Button
        Private btnLogout As Button

        ' Sidebar Navigation Items
        Private navButtons As New List(Of Button)()
        Private navNames As New List(Of String)()
        Private activeNavIndex As Integer = 0
        Private activeViewRetry As Action = Nothing
        Private sidebarCollapsed As Boolean = False
        Private tipNav As ToolTip
        Private pnlBrand As Panel

        ' Views (Panels)
        Private viewDashboard As Panel
        Private viewRegistry As Panel
        Private viewDirectives As Panel
        Private viewSearch As Panel
        Private viewAdmin As Panel
        Private viewAudit As Panel
        Private viewPortalIntake As Panel

        ' Portal Intake Controls
        Private dgvPortalQueue As DataGridView
        Private btnPortalRefresh As Button
        Private btnPortalImportSelected As Button
        Private lblPortalStatus As Label
        Private lblPortalWatermark As Label
        Private _portalSubmissions As New List(Of PortalSubmission)()

        ' Dashboard Controls
        Private lblStatTotalDocs As Label
        Private lblStatDirectives As Label
        Private lblStatActiveRoute As Label
        Private lblStatVaultStorage As Label
        Private cmbDashSection As ComboBox
        Private cmbDashCategory As ComboBox
        Private btnDashResetFilters As Button
        Private dgvDashRecent As DataGridView
        Private lblDashWatermark As Label

        ' Registry Controls
        Private txtTitle As TextBox
        Private cmbDocType As ComboBox
        Private cmbFlowDirection As ComboBox
        Private cmbDeadlinePreset As ComboBox
        Private dtpDeadline As DateTimePicker
        Private txtOrigin As TextBox
        Private txtDest As TextBox
        Private txtCabinet As TextBox
        Private txtShelf As TextBox
        Private txtBox As TextBox
        Private txtGDrive As TextBox
        Private btnAttachScan As Button
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
        Private cmbNewUserSection As ComboBox
        Private lblSectionNotice As Label
        Private chkCanRoute As CheckBox
        Private chkCanMove As CheckBox
        Private chkCanSoftCopy As CheckBox
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

            ' Run initial SQL sync before opening login dialog so badges enrolled on other desks are active.
            AddHandler Me.Shown, Async Sub()
                                     Await RunSqlSyncAsync()
                                     ShowRFIDLoginDialog()
                                 End Sub
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
            SetupPortalIntakeView()

            ' Buffer every container once so maximize/restore repaints atomically instead of
            ' flashing through half-painted panels and borders.
            UiBuffering.EnableDeep(Me)

            SwitchNavView(0)

            If AppSettings.Instance.PortalSettings.PortalEnabled Then
                OnPortalPollTick()
            End If
        End Sub

        Private Sub SetupFooterStatusStrip()
            stsFooter = New StatusStrip With {
                .Dock = DockStyle.Bottom,
                .BackColor = CivicCalmTheme.ColorWell,
                .ForeColor = CivicCalmTheme.ColorInk,
                .Font = CivicCalmTheme.FontMicrocopy,
                .SizingGrip = False
            }

            Dim dbBadge = If(Program.IsDatabaseConnected, "[ DB: SQL Server Connected ]", "[ DB: Offline Embedded Cache ]")
            lblStatusMessage = New ToolStripStatusLabel With {
                .Text = "Ready " & dbBadge,
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

            tmrClock = New Timer With {
                .Interval = 1000,
                .Enabled = True
            }
            AddHandler tmrClock.Tick, Sub()
                                          If lblStatusClock IsNot Nothing AndAlso Not Me.IsDisposed Then
                                              lblStatusClock.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                                          End If
                                      End Sub

            ' ponytail: 5s poll. Two or three workstations on one LAN do not need faster, and a
            ' shorter interval is one constant away if the desk ever asks for it.
            Const SqlSyncIntervalMs As Integer = 5000
            tmrSqlSync = New Timer With {
                .Interval = SqlSyncIntervalMs,
                .Enabled = True
            }
            AddHandler tmrSqlSync.Tick, Sub() OnSqlSyncTick()

            If AppSettings.Instance.PortalSettings.PortalEnabled Then
                Dim pollSec = Math.Max(10, AppSettings.Instance.PortalSettings.PollIntervalSeconds)
                tmrPortalPoll = New Timer With {
                    .Interval = pollSec * 1000,
                    .Enabled = True
                }
                AddHandler tmrPortalPoll.Tick, Sub() OnPortalPollTick()
            End If
        End Sub

        Private Sub OnPortalPollTick()
            If Me.IsDisposed OrElse Not Me.IsHandleCreated Then Return
            If AppStartup.PortalBridgeClient Is Nothing OrElse Not AppSettings.Instance.PortalSettings.PortalEnabled Then Return
            Task.Run(Async Function()
                         Try
                             Dim items = Await AppStartup.PortalBridgeClient.FetchExternalQueueAsync().ConfigureAwait(False)
                             If Me.IsDisposed OrElse Not Me.IsHandleCreated Then Return
                             Me.BeginInvoke(Sub()
                                                If Not Me.IsDisposed AndAlso Me.IsHandleCreated Then
                                                    UpdatePortalIntakeBadgeAndQueue(items)
                                                End If
                                            End Sub)
                         Catch ex As Exception
                             System.Diagnostics.Trace.WriteLine("Portal polling exception: " & ex.Message)
                         End Try
                     End Function)
        End Sub

        ''' <summary>
        ''' Pulls SQL Server into the cache on a background thread, then applies the snapshot
        ''' on the UI thread (Await resumes here) so bound grids refresh on one layout pass.
        ''' A disconnected or unreachable host costs one pointer flip and nothing else.
        ''' </summary>
        Private Async Function RunSqlSyncAsync() As Task
            If Me.IsDisposed OrElse Not Me.IsHandleCreated Then Return
            If Not Program.IsDatabaseConnected OrElse Program.Coordinator Is Nothing Then Return
            If System.Threading.Interlocked.CompareExchange(_sqlSyncInFlight, 1, 0) <> 0 Then Return
            Try
                Await Task.Run(Sub() Program.Coordinator.SyncOfflineOutbox())
                Dim snapshot = Await Task.Run(Function() Program.Coordinator.BuildSqlSnapshot())
                If snapshot Is Nothing OrElse Me.IsDisposed Then Return
                ApplySqlMirror(snapshot)
            Catch ex As Exception
                System.Diagnostics.Trace.WriteLine("SQL mirror sync failed: " & ex.Message)
            Finally
                System.Threading.Interlocked.Exchange(_sqlSyncInFlight, 0)
            End Try
        End Function

        Private Async Sub OnSqlSyncTick()
            Await RunSqlSyncAsync()
        End Sub

        Private Sub ApplySqlMirror(snapshot As Dictionary(Of String, DataTable))
            EmbeddedDB.ApplySnapshots(snapshot, persist:=False)
            RebindCurrentUser()
            If snapshot IsNot Nothing AndAlso snapshot.ContainsKey("Users") Then
                PopulateStaffDropdowns()
            End If
            RefreshActiveTabGrid()
        End Sub

        ''' <summary>
        ''' The session row is captured at login, while the mirror moves Users rows from
        ''' cache-local ids onto SQL identities. Re-point it at the mirrored row for the same
        ''' badge so UserID stays the SQL identity every write path puts into a foreign key.
        ''' When the mirrored table no longer carries the badge (offline enrolment that SQL
        ''' never saw, or a disabled account) the held row is detached and every later read
        ''' throws RowNotInTableException, so the session drops to the logged-out state that
        ''' every CurrentUser call site already checks for.
        ''' </summary>
        Private Sub RebindCurrentUser()
            If CurrentUser Is Nothing Then Return
            If Not CurrentUser.Table.Columns.Contains("RFID_UID") Then Return
            Dim uid = CurrentUser("RFID_UID").ToString()
            If String.IsNullOrWhiteSpace(uid) Then Return
            Dim match = EmbeddedDB.DataSet.Tables("Users").Select(String.Format("RFID_UID = '{0}' AND IsActive = True", uid.Replace("'", "''")))
            If match.Length > 0 Then
                CurrentUser = match(0)
            Else
                System.Diagnostics.Trace.WriteLine("Session badge " & uid & " is not in the mirrored user list; ending the session view.")
                CurrentUser = Nothing
            End If
        End Sub

        Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
            ' Flush deferred embedded-store writes (mutations batch via EmbeddedDB.MarkDirty).
            EmbeddedDB.Save()
            If tmrPortalPoll IsNot Nothing Then
                tmrPortalPoll.Stop()
                tmrPortalPoll.Dispose()
                tmrPortalPoll = Nothing
            End If
            If tmrSqlSync IsNot Nothing Then
                tmrSqlSync.Stop()
                tmrSqlSync.Dispose()
                tmrSqlSync = Nothing
            End If
            If tmrClock IsNot Nothing Then
                tmrClock.Stop()
                tmrClock.Dispose()
                tmrClock = Nothing
            End If
            MyBase.OnFormClosing(e)
        End Sub

        Private Sub SetupSidebar()
            pnlSidebar = New Panel With {
                .Dock = DockStyle.Left,
                .Width = 246,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(12)
            }

            pnlBrand = New Panel With {
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

            tipNav = New ToolTip()

            Dim btnToggle As New Button With {
                .Name = "btnNav_Toggle",
                .Text = "«  Collapse",
                .Size = New Size(218, 30),
                .Margin = New Padding(0, 0, 0, 10),
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontMicrocopy,
                .Cursor = Cursors.Hand,
                .TextAlign = ContentAlignment.MiddleLeft,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .BackColor = CivicCalmTheme.ColorWell
            }
            btnToggle.FlatAppearance.BorderSize = 1
            btnToggle.FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder
            btnToggle.FlatAppearance.MouseOverBackColor = CivicCalmTheme.ColorWell
            AddHandler btnToggle.Click, Sub() ToggleSidebar(btnToggle)
            pnlNavStack.Controls.Add(btnToggle)

            Dim navItems As New List(Of String) From {"Dashboard", "Document Registry", "SG Directives", "Search & Storage", "User & RFID Admin", "Audit Trail", "Portal Intake"}

            For i As Integer = 0 To navItems.Count - 1
                Dim idx As Integer = i
                navNames.Add(navItems(i))
                Dim btn As New Button With {
                    .Name = "btnNav_" & navItems(i).Replace(" "c, "_"c),
                    .Text = "  " & navItems(i),
                    .Size = New Size(218, 44),
                    .Margin = New Padding(0, 0, 0, 6),
                    .FlatStyle = FlatStyle.Flat,
                    .Font = CivicCalmTheme.FontBody,
                    .TextAlign = ContentAlignment.MiddleLeft,
                    .Cursor = Cursors.Hand,
                    .UseMnemonic = False,
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
            ' Lowest dock index docks last, so the Fill nav stack is measured after the
            ' Top brand block and cannot be overlapped by it.
            pnlNavStack.BringToFront()
            Me.Controls.Add(pnlSidebar)
        End Sub

        ' DESIGN.md variance exception: the docked sidebar collapses to a 56px rail; only text
        ' is dropped, order and behaviour are identical, and hover tooltips identify the views.
        Private Sub ToggleSidebar(btnToggle As Button)
            sidebarCollapsed = Not sidebarCollapsed
            pnlSidebar.Width = If(sidebarCollapsed, 56, 246)
            pnlBrand.Visible = Not sidebarCollapsed
            For i As Integer = 0 To navButtons.Count - 1
                If sidebarCollapsed Then
                    navButtons(i).Text = ""
                    navButtons(i).Size = New Size(32, 44)
                    Dim tip = navNames(i)
                    If i = 6 AndAlso _portalSubmissions IsNot Nothing AndAlso _portalSubmissions.Count > 0 Then
                        tip = $"Portal Intake ({_portalSubmissions.Count} pending)"
                    End If
                    tipNav.SetToolTip(navButtons(i), tip)
                Else
                    Dim labelText = navNames(i)
                    If i = 6 AndAlso _portalSubmissions IsNot Nothing AndAlso _portalSubmissions.Count > 0 Then
                        labelText = $"Portal Intake ({_portalSubmissions.Count})"
                    End If
                    navButtons(i).Text = "  " & labelText
                    navButtons(i).Size = New Size(218, 44)
                    tipNav.SetToolTip(navButtons(i), Nothing)
                End If
            Next
            btnToggle.Text = If(sidebarCollapsed, "»", "«  Collapse")
            btnToggle.Size = If(sidebarCollapsed, New Size(32, 30), New Size(218, 30))
            btnToggle.TextAlign = If(sidebarCollapsed, ContentAlignment.MiddleCenter, ContentAlignment.MiddleLeft)
            tipNav.SetToolTip(btnToggle, If(sidebarCollapsed, "Expand navigation", Nothing))
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

            ' Controls.Clear() disposes what it removes, and these view panels are built once, so
            ' remove them without disposing or the next visit re-adds a dead control. The swap and
            ' the grid refresh are wrapped so the whole switch settles in a single layout pass.
            pnlContent.SuspendLayout()
            Try
                While pnlContent.Controls.Count > 0
                    pnlContent.Controls.RemoveAt(0)
                End While
                Select Case index
                    Case 0 : pnlContent.Controls.Add(viewDashboard)
                    Case 1 : pnlContent.Controls.Add(viewRegistry)
                    Case 2 : pnlContent.Controls.Add(viewDirectives)
                    Case 3 : pnlContent.Controls.Add(viewSearch)
                    Case 4 : pnlContent.Controls.Add(viewAdmin)
                    Case 5 : pnlContent.Controls.Add(viewAudit)
                    Case 6 : If viewPortalIntake IsNot Nothing Then pnlContent.Controls.Add(viewPortalIntake)
                End Select

                RefreshActiveTabGrid()
            Finally
                pnlContent.ResumeLayout(True)
            End Try
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
                .Font = CivicCalmTheme.FontFieldLabel,
                .ForeColor = CivicCalmTheme.ColorDanger,
                .AutoSize = True
            }

            pnlTitles.Controls.AddRange(New Control() {lblTitle, lblUserBadge})

            Dim pnlAuthActions As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.RightToLeft,
                .WrapContents = False,
                .BackColor = CivicCalmTheme.ColorSurface
            }

            btnLogout = New Button With {
                .Text = "&Logout",
                .Size = New Size(90, 36),
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontBody,
                .BackColor = CivicCalmTheme.ColorWell,
                .ForeColor = CivicCalmTheme.ColorInk,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(8, 4, 0, 0)
            }
            btnLogout.FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder
            AddHandler btnLogout.Click, Sub() AuthenticateUser("")

            btnScanRFID = New Button With {
                .Text = "&Tap RFID Smart Card",
                .Size = New Size(180, 36),
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .BackColor = CivicCalmTheme.ColorPrimary,
                .ForeColor = Color.White,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(0, 4, 0, 0)
            }
            btnScanRFID.FlatAppearance.BorderSize = 0
            AddHandler btnScanRFID.Click, Sub() ShowRFIDLoginDialog()

            pnlAuthActions.Controls.AddRange(New Control() {btnLogout, btnScanRFID})

            tblHeader.Controls.Add(pnlTitles, 0, 0)
            tblHeader.Controls.Add(pnlAuthActions, 1, 0)
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
                lblUserBadge.Text = "[ RFID Logged Out: Access Restricted ]"
                lblUserBadge.ForeColor = CivicCalmTheme.ColorDanger
                lblStatusMessage.Text = "User logged out."
                If cmbDashSection IsNot Nothing Then
                    cmbDashSection.Enabled = True
                    cmbDashSection.SelectedIndex = 0
                End If
            Else
                Dim maskedUid As String = If(uid.Length > 4, "****" & uid.Substring(uid.Length - 4), uid)
                Dim user = EmbeddedDB.AuthenticateRFID(uid)
                If user IsNot Nothing Then
                    CurrentUser = user
                    Dim role = user("Role").ToString()
                    Dim office = If(user.Table.Columns.Contains("Office") AndAlso Not IsDBNull(user("Office")), user("Office").ToString(), "")
                    Dim effectiveDesk = If(Not String.IsNullOrWhiteSpace(office), office, role)
                    Dim isGlobal As Boolean = (role = "Secretary-General" OrElse role = "System Administrator" OrElse role = "OSG Chief")
                    Dim displayRole = If(Not String.IsNullOrWhiteSpace(office) AndAlso Not office.Equals(role, StringComparison.OrdinalIgnoreCase), String.Format("{0} - {1}", role.ToUpperInvariant(), office.ToUpperInvariant()), role.ToUpperInvariant())
                    lblUserBadge.Text = String.Format("AUTHENTICATED: {0} [{1}] : {2}", user("FullName").ToString().ToUpperInvariant(), displayRole, If(isGlobal, "GLOBAL ACCESS", "SECTION DESK VIEW"))
                    lblUserBadge.ForeColor = CivicCalmTheme.ColorPrimary
                    lblStatusMessage.Text = "Authenticated: " & user("FullName").ToString()
                    EmbeddedDB.LogAudit(user("FullName").ToString(), "RFID Badge Tap Authenticated [Card: " & maskedUid & "]")

                    If cmbDashSection IsNot Nothing Then
                        If isGlobal Then
                            cmbDashSection.Enabled = True
                        Else
                            If effectiveDesk.EndsWith("Section", StringComparison.OrdinalIgnoreCase) OrElse effectiveDesk.Equals("Secretariat", StringComparison.OrdinalIgnoreCase) Then
                                Dim idx = cmbDashSection.Items.IndexOf(effectiveDesk)
                                If idx >= 0 Then
                                    cmbDashSection.SelectedIndex = idx
                                End If
                            End If
                            cmbDashSection.Enabled = False
                        End If
                    End If
                Else
                    If EmbeddedDB.IsTerminalLockedOut() Then
                        lblStatusMessage.Text = "Security Lockout: Terminal locked for 5 minutes due to 5 consecutive failed card reads."
                        MessageBox.Show("Terminal has been temporarily locked out due to 5 consecutive failed RFID smart card badge reads. Please notify the System Administrator or wait 5 minutes.", "Security Lockout", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Else
                        lblStatusMessage.Text = "Access Denied: Unrecognized RFID card [" & maskedUid & "]"
                        MessageBox.Show("Unrecognized RFID Smart Card Badge UID: " & maskedUid, "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    End If
                End If
            End If
            PopulateStaffDropdowns()
            RefreshActiveTabGrid()
        End Sub

        Public Sub PopulateStaffDropdowns()
            Dim prevStaff = If(cmbAssignedStaff.SelectedItem, "").ToString()
            Dim prevDirAssign = If(cmbDirAssign.SelectedItem, "").ToString()

            cmbAssignedStaff.BeginUpdate()
            cmbDirAssign.BeginUpdate()
            Try
                cmbAssignedStaff.Items.Clear()
                cmbDirAssign.Items.Clear()
                For Each row As DataRow In EmbeddedDB.DataSet.Tables("Users").Rows
                    Dim name = row("FullName").ToString()
                    cmbAssignedStaff.Items.Add(name)
                    cmbDirAssign.Items.Add(name)
                Next

                Dim staffIdx = If(Not String.IsNullOrEmpty(prevStaff), cmbAssignedStaff.Items.IndexOf(prevStaff), -1)
                If staffIdx >= 0 Then
                    cmbAssignedStaff.SelectedIndex = staffIdx
                ElseIf cmbAssignedStaff.Items.Count > 0 Then
                    cmbAssignedStaff.SelectedIndex = 0
                End If

                Dim dirIdx = If(Not String.IsNullOrEmpty(prevDirAssign), cmbDirAssign.Items.IndexOf(prevDirAssign), -1)
                If dirIdx >= 0 Then
                    cmbDirAssign.SelectedIndex = dirIdx
                ElseIf cmbDirAssign.Items.Count > 0 Then
                    cmbDirAssign.SelectedIndex = 0
                End If
            Finally
                cmbAssignedStaff.EndUpdate()
                cmbDirAssign.EndUpdate()
            End Try
        End Sub

        Private Sub ApplyCardBorder(pnl As Panel)
            AddHandler pnl.Resize, Sub(s As Object, e As EventArgs) pnl.Invalidate()
            AddHandler pnl.Paint, Sub(s As Object, pe As PaintEventArgs)
                                      ControlPaint.DrawBorder(pe.Graphics, pnl.ClientRectangle, CivicCalmTheme.ColorBorder, ButtonBorderStyle.Solid)
                                  End Sub
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
            ' While the label doubles as the DESIGN 4.3 error banner, clicking it reruns the
            ' registered retry action. Outside an error no retry is registered, so the click
            ' does nothing and the banner stays non-interactive.
            AddHandler lbl.Click, Sub() TryRetryActiveView()
            parent.Controls.Add(lbl)
            Return lbl
        End Function

        Public Sub ReportStatus(message As String)
            lblStatusMessage.Text = message
        End Sub

        Private Sub TryRetryActiveView()
            If activeViewRetry Is Nothing Then Return
            lblStatusMessage.Text = "Retrying the failed load..."
            activeViewRetry()
        End Sub

        Private Function Dpi(v As Single) As Single
            Return v * DeviceDpi / 96.0F
        End Function

        Public Sub OpenSelectedDocumentDetail(dgv As DataGridView)
            If dgv Is Nothing OrElse dgv.CurrentRow Is Nothing OrElse dgv.CurrentRow.Cells("DocumentID").Value Is Nothing OrElse IsDBNull(dgv.CurrentRow.Cells("DocumentID").Value) Then
                lblStatusMessage.Text = "Please select a document row from the grid first."
                Return
            End If

            Dim docId = CInt(dgv.CurrentRow.Cells("DocumentID").Value)
            Using dlg As New FormDocumentDetail(docId, Me)
                dlg.ShowDialog(Me)
            End Using
            RefreshActiveTabGrid()
        End Sub

        Public Sub PrintSelectedDocumentRoutingSlip(dgv As DataGridView)
            If dgv Is Nothing OrElse dgv.CurrentRow Is Nothing OrElse dgv.CurrentRow.Cells("DocumentID").Value Is Nothing OrElse IsDBNull(dgv.CurrentRow.Cells("DocumentID").Value) Then
                lblStatusMessage.Text = "Please select a document row from the grid first to print routing slip."
                Return
            End If

            Dim docId = CInt(dgv.CurrentRow.Cells("DocumentID").Value)
            Dim docObj = EmbeddedDB.GetDocumentByID(docId)
            If docObj Is Nothing Then Return

            Dim logs = EmbeddedDB.GetRoutingLogsForDocument(docId)
            Dim directives = EmbeddedDB.GetDirectivesForDocument(docId)

            RoutingSlipPrintService.ShowPreview(docObj, logs, directives, Me)
        End Sub

        Public Sub RefreshActiveTabGrid()
            Dim selSec As String = If(cmbDashSection IsNot Nothing AndAlso cmbDashSection.SelectedItem IsNot Nothing, cmbDashSection.SelectedItem.ToString(), "")
            Dim selCat As String = If(cmbDashCategory IsNot Nothing AndAlso cmbDashCategory.SelectedItem IsNot Nothing, cmbDashCategory.SelectedItem.ToString(), "")
            Dim visibleDocs = EmbeddedDB.GetVisibleDocuments(CurrentUser, selSec, selCat)

            Select Case activeNavIndex
                Case 0 ' Dashboard View (stats and recent grid only refresh while the dashboard is on screen)
                    Dim allDocs = EmbeddedDB.DataSet.Tables("Documents")
                    Dim forReviewCount As Integer = 0
                    Dim forRevisionCount As Integer = 0
                    Dim approvedReleasedCount As Integer = 0
                    For Each row As DataRow In allDocs.Rows
                        Dim st = row("CurrentStatus").ToString()
                        If st.IndexOf("REVIEW", StringComparison.OrdinalIgnoreCase) >= 0 Then forReviewCount += 1
                        If st.IndexOf("REVISION", StringComparison.OrdinalIgnoreCase) >= 0 Then forRevisionCount += 1
                        If st.Equals("APPROVED", StringComparison.OrdinalIgnoreCase) OrElse st.Equals("RELEASED", StringComparison.OrdinalIgnoreCase) Then approvedReleasedCount += 1
                    Next

                    lblStatTotalDocs.Text = allDocs.Rows.Count.ToString()
                    lblStatDirectives.Text = forReviewCount.ToString()
                    lblStatActiveRoute.Text = forRevisionCount.ToString()
                    lblStatVaultStorage.Text = approvedReleasedCount.ToString()
                    dgvDashRecent.DataSource = visibleDocs
                    DataGridStyler.FormatDocumentColumns(dgvDashRecent)
                    If dgvDashRecent.Rows.Count = 0 Then
                        DataGridStyler.SetEmptyState(dgvDashRecent, lblDashWatermark)
                    Else
                        DataGridStyler.SetPopulatedState(dgvDashRecent, lblDashWatermark)
                    End If
                Case 1 ' Registry View
                    dgvRegistry.DataSource = visibleDocs
                    DataGridStyler.FormatDocumentColumns(dgvRegistry)
                    If dgvRegistry.Rows.Count = 0 Then
                        DataGridStyler.SetEmptyState(dgvRegistry, lblRegistryWatermark)
                    Else
                        DataGridStyler.SetPopulatedState(dgvRegistry, lblRegistryWatermark)
                    End If
                    lblStatusCount.Text = dgvRegistry.Rows.Count.ToString() & " Documents"
                Case 2 ' Directives View
                    Dim previousDocId As Integer = 0
                    If cmbDirDocs.SelectedItem IsNot Nothing Then
                        Dim selectedStr = cmbDirDocs.SelectedItem.ToString()
                        Dim match = System.Text.RegularExpressions.Regex.Match(selectedStr, "^ID (\d+):")
                        If match.Success Then Integer.TryParse(match.Groups(1).Value, previousDocId)
                    End If

                    cmbDirDocs.BeginUpdate()
                    Try
                        cmbDirDocs.Items.Clear()
                        Dim restoreIdx As Integer = -1
                        Dim currentIdx As Integer = 0
                        For Each drv As DataRowView In visibleDocs
                            Dim docId = Convert.ToInt32(drv("DocumentID"))
                            cmbDirDocs.Items.Add(String.Format("ID {0}: [{1}] {2}", docId, drv("DocCode"), drv("Title")))
                            If docId = previousDocId Then restoreIdx = currentIdx
                            currentIdx += 1
                        Next
                        If restoreIdx >= 0 Then
                            cmbDirDocs.SelectedIndex = restoreIdx
                        ElseIf cmbDirDocs.Items.Count > 0 Then
                            cmbDirDocs.SelectedIndex = 0
                        End If
                    Finally
                        cmbDirDocs.EndUpdate()
                    End Try
                    Dim dtDirectives = EmbeddedDB.DataSet.Tables("Directives")
                    dgvDirectives.DataSource = dtDirectives
                    DataGridStyler.FormatDirectiveColumns(dgvDirectives)
                    If dgvDirectives.Rows.Count = 0 Then
                        DataGridStyler.SetEmptyState(dgvDirectives, lblDirectivesWatermark, "No active action directives found.")
                    Else
                        DataGridStyler.SetPopulatedState(dgvDirectives, lblDirectivesWatermark)
                    End If
                    lblStatusCount.Text = dgvDirectives.Rows.Count.ToString() & " Directives"
                Case 3 ' Search View
                    dgvSearch.DataSource = visibleDocs
                    DataGridStyler.FormatDocumentColumns(dgvSearch)
                    If dgvSearch.Rows.Count = 0 Then
                        DataGridStyler.SetEmptyState(dgvSearch, lblSearchWatermark)
                    Else
                        DataGridStyler.SetPopulatedState(dgvSearch, lblSearchWatermark)
                    End If
                    lblStatusCount.Text = dgvSearch.Rows.Count.ToString() & " Documents"
                Case 4 ' Admin View
                    Dim dtUsers = EmbeddedDB.DataSet.Tables("Users")
                    dgvUsers.DataSource = dtUsers
                    DataGridStyler.FormatUserColumns(dgvUsers)
                    If dgvUsers.Rows.Count = 0 Then
                        DataGridStyler.SetEmptyState(dgvUsers, lblUsersWatermark, "No users registered.")
                    Else
                        DataGridStyler.SetPopulatedState(dgvUsers, lblUsersWatermark)
                    End If
                    lblStatusCount.Text = dgvUsers.Rows.Count.ToString() & " Users"
                Case 5 ' Audit View
                    If txtAuditSearch Is Nothing OrElse String.IsNullOrWhiteSpace(txtAuditSearch.Text) Then
                        Dim dtAudit = EmbeddedDB.DataSet.Tables("AuditTrail")
                        dgvAudit.DataSource = dtAudit
                        DataGridStyler.FormatAuditColumns(dgvAudit)
                        If dgvAudit.Rows.Count = 0 Then
                            DataGridStyler.SetEmptyState(dgvAudit, lblAuditWatermark, "No audit trail events recorded.")
                        Else
                            DataGridStyler.SetPopulatedState(dgvAudit, lblAuditWatermark)
                        End If
                        lblStatusCount.Text = dgvAudit.Rows.Count.ToString() & " Audit Events"
                    End If
                Case 6 ' Portal Intake View
                    RefreshPortalQueue()
            End Select

            lblStatusClock.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        End Sub

        Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
            If keyData = (Keys.Alt Or Keys.C) Then
                ClearActiveViewFilters()
                Return True
            End If
            If keyData = (Keys.Alt Or Keys.R) Then
                TryRetryActiveView()
                Return True
            End If
            Return MyBase.ProcessCmdKey(msg, keyData)
        End Function

        Private Sub ClearActiveViewFilters()
            Select Case activeNavIndex
                Case 0, 1
                    btnDashResetFilters.PerformClick()
                Case 3
                    btnClearSearch.PerformClick()
                Case 5
                    btnAuditReset.PerformClick()
                Case Else
                    lblStatusMessage.Text = "This view has no filters to clear."
                    Return
            End Select
            lblStatusMessage.Text = "Filters cleared for the current view."
        End Sub

        Protected Overrides Sub Dispose(disposing As Boolean)
            If disposing Then
                If tmrClock IsNot Nothing Then
                    tmrClock.Stop()
                    tmrClock.Dispose()
                    tmrClock = Nothing
                End If
            End If
            MyBase.Dispose(disposing)
        End Sub
    End Class
End Namespace
