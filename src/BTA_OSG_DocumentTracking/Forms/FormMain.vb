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
        Private lblDbState As ToolStripStatusLabel
        Private epValidation As ErrorProvider
        Private tmrClock As Timer
        Private tmrPortalPoll As Timer
        Private tmrSqlSync As Timer

        ' One SQL snapshot pull at a time: a slow host skips ticks instead of queueing them.
        Private _sqlSyncInFlight As Integer

        ' Bind-once state: grids keep their DataView across sync ticks and the in-place
        ' mirror merge updates them through row events instead of full rebinds. The view is
        ' recreated only when the session or the dashboard filters change.
        Private _activeDocsView As DataView
        Private _activeDocsViewKey As String
        Private _auditView As DataView
        Private tmrResizeRefresh As Timer

        ' Header Controls
        Private lblTitle As Label
        Private lblUserBadge As Label
        Private btnCheckUpdate As Button
        Private btnScanRFID As Button
        Private btnLogout As Button
        Private btnRevokeCard As Button

        ' Sidebar Navigation Items
        Private navButtons As New List(Of Button)()
        Private navNames As New List(Of String)()
        Private navIconNames As New List(Of String)()
        Private activeNavIndex As Integer = 0
        Private activeViewRetry As Action = Nothing
        Private tipNav As ToolTip
        Private pnlBrand As Panel

        ' Views (Panels)
        Private viewDashboard As Panel
        Private viewAnalytics As Panel
        Private viewRegistry As Panel
        Private viewDirectives As Panel
        Private viewSearch As Panel
        Private viewAdmin As Panel
        Private viewAudit As Panel
        Private viewPortalIntake As Panel

        ' Portal Intake Controls
        Private dgvPortalQueue As DataGridView
        Private btnOpenPortal As Button
        Private btnStartTunnel As Button
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
        Private lblDeadlineHint As Label
        Private cmbOrigin As ComboBox
        Private cmbDest As ComboBox
        Private cmbCabinet As ComboBox
        Private cmbShelf As ComboBox
        Private cmbBox As ComboBox
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
        Private btnUnlockUser As Button

        ' Audit Controls
        Private txtAuditSearch As TextBox
        Private btnAuditFilter As Button
        Private btnAuditReset As Button
        Private btnAuditVerify As Button
        Private dgvAudit As DataGridView
        Private lblAuditWatermark As Label

        ' Workstation Sync Status (Admin tab)
        Private dgvSeats As DataGridView
        Private lblSeatsWatermark As Label

        ' One draw per process: the whole point is that simultaneous morning boots do not
        ' land on the same tick, so a fresh Random per instance is enough.
        Private ReadOnly _sqlSyncJitterMs As Integer = New Random().Next(0, 3000)

        Public Sub New()
            EmbeddedDB.Initialize()
            InitializeUI()
            PopulateStaffDropdowns()

            ' Data-layer failures must reach the operator's banner, not just the Trace log
            ' of a WinExe: SQL fallbacks, replay errors, and cache save failures.
            Program.OperatorWarningSink = AddressOf ShowOperatorWarningThreadSafe
            AddHandler EmbeddedDB.PersistenceFailed, AddressOf OnCacheSaveFailed
            SetStyle(ControlStyles.OptimizedDoubleBuffer Or ControlStyles.AllPaintingInWmPaint, True)
            UpdateStyles()

            ' Sync only when the startup probe found the host live; the offline re-probe belongs
            ' to the sync timer tick, not to a 5s stall in front of the login dialog. The grid
            ' refresh runs here so the dashboard balances against the real client size.
            AddHandler Me.Shown, Async Sub()
                                     ' A dropped settings layer (unreadable DPAPI blob, malformed JSON)
                                     ' fail-opens to the shipped defaults: the operator must be told
                                     ' or the seat can silently point at the wrong server.
                                     If AppSettings.Instance.IgnoredConfigLayers.Count > 0 Then
                                         ShowOperatorWarning("Warning: " & AppSettings.Instance.IgnoredConfigLayers(0) & ". Saved settings were ignored; check Station Setup.")
                                     End If
                                     If Program.IsDatabaseConnected Then Await RunSqlSyncAsync()
                                     RefreshActiveTabGrid()
                                     If AppSettings.Instance.PortalSettings.PortalEnabled Then OnPortalPollTick()
                                     ShowRFIDLoginDialog()
                                 End Sub
        End Sub

        Private Sub InitializeUI()
            AppAssets.ApplyFormIcon(Me)
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
            SetupAnalyticsView()
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
                Task.Run(Async Function()
                             Await PortalServerManager.EnsureRunningAsync(AddressOf SetPortalStatusThreadSafe).ConfigureAwait(False)
                         End Function)
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

            ' The DB badge lives in its own slot so the clock tick can restate it when the
            ' sync poller flips the connection mid-session, without clobbering operator
            ' warnings that use the message slot.
            lblDbState = New ToolStripStatusLabel With {
                .Text = dbBadge,
                .BorderSides = ToolStripStatusLabelBorderSides.Left,
                .BorderStyle = Border3DStyle.Etched,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .Padding = New Padding(8, 0, 8, 0)
            }

            Dim currentVer = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version
            Dim verText = If(currentVer IsNot Nothing, $"v{AppUpdateService.FormatVersion(currentVer)}", "v?")
            Dim lblVersion = New ToolStripStatusLabel With {
                .Text = $"{verText} (What's New)",
                .BorderSides = ToolStripStatusLabelBorderSides.Left,
                .BorderStyle = Border3DStyle.Etched,
                .ForeColor = CivicCalmTheme.ColorPrimary,
                .Padding = New Padding(8, 0, 8, 0),
                .IsLink = True,
                .LinkBehavior = LinkBehavior.HoverUnderline
            }
            AddHandler lblVersion.Click, Async Sub() Await AppUpdateService.CheckAndApplyUpdateAsync(Me, True)

            stsFooter.Items.AddRange(New ToolStripItem() {lblStatusMessage, lblStatusCount, lblDbState, lblVersion, lblStatusClock})
            Me.Controls.Add(stsFooter)

            tmrClock = New Timer With {
                .Interval = 1000,
                .Enabled = True
            }
            AddHandler tmrClock.Tick, Sub()
                                          If lblStatusClock IsNot Nothing AndAlso Not Me.IsDisposed Then
                                              lblStatusClock.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                                          End If
                                          If lblDbState IsNot Nothing AndAlso Not Me.IsDisposed Then
                                              lblDbState.Text = If(Program.IsDatabaseConnected, "[ DB: SQL Server Connected ]", "[ DB: Offline Embedded Cache ]")
                                          End If
                                          EnforceSessionTimeout()
                                      End Sub

            ' Multi-workstation sync polling: the configured interval (floor 5s) plus a
            ' one-time jitter so seats booted at the same time do not poll in one wave.
            tmrSqlSync = New Timer With {
                .Interval = AppSettings.Instance.DatabaseSettings.EffectiveSyncIntervalSeconds() * 1000 + _sqlSyncJitterMs,
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
                             ' The badge would otherwise go stale with no on-screen hint, so
                             ' surface the failure the same way RefreshPortalQueue does.
                             If Me.IsDisposed OrElse Not Me.IsHandleCreated Then Return
                             Me.BeginInvoke(Sub()
                                                If Not Me.IsDisposed AndAlso Me.IsHandleCreated AndAlso lblPortalStatus IsNot Nothing Then
                                                    lblPortalStatus.Text = "Portal Bridge: Offline / Unreachable"
                                                    lblStatusMessage.Text = "Portal polling failed: " & ex.Message
                                                End If
                                            End Sub)
                         End Try
                     End Function)
        End Sub

        ''' <summary>
        ''' Pulls SQL Server into the cache on a background thread, then applies the snapshot
        ''' on the UI thread (Await resumes here) so bound grids refresh on one layout pass.
        ''' A disconnected or unreachable host costs one pointer flip and nothing else.
        ''' The tick also runs while the probe reports offline: BuildSqlSnapshot's connect
        ''' attempt is the re-probe that lets a workstation recover after a startup blip.
        ''' </summary>
        Private Async Function RunSqlSyncAsync() As Task
            If Me.IsDisposed OrElse Not Me.IsHandleCreated Then Return
            If Program.Coordinator Is Nothing Then Return
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
            Dim changedTables = EmbeddedDB.ApplySnapshots(snapshot, persist:=False)
            RebindCurrentUser()
            ' The workstation heartbeat changes on every pull by design, so it can never
            ' count as a change; otherwise every sync tick would pay for a full refresh.
            ' A heartbeat-only tick may still matter to the Admin tab, whose seat grid
            ' recomputes staleness at bind time.
            changedTables.Remove("Heartbeat")
            If changedTables.Count = 0 Then
                If activeNavIndex = 5 Then RefreshActiveTabGrid()
                Return
            End If
            If changedTables.Contains("Users") AndAlso snapshot.ContainsKey("Users") Then
                PopulateStaffDropdowns()
            End If
            RefreshActiveTabGrid()
        End Sub

        ''' <summary>
        ''' The configured inactivity timeout is enforced here: the clock timer checks real
        ''' user input (UiActivityMonitor) every second and ends the session when the desk
        ''' has been idle past the limit. Until this existed the TimeoutMinutes setting had
        ''' no caller and sessions lived until manual logout.
        ''' </summary>
        Private Sub EnforceSessionTimeout()
            If CurrentUser Is Nothing OrElse Me.IsDisposed Then Return
            Dim timeoutMinutes = AppSettings.Instance.SessionSettings.TimeoutMinutes
            If timeoutMinutes <= 0 Then Return
            If UiActivityMonitor.Instance.IdleSeconds < timeoutMinutes * 60 Then Return

            Dim who As String = "Unknown Staff"
            Try
                who = CurrentUser("FullName").ToString()
            Catch
            End Try
            EmbeddedDB.LogAudit(who, String.Format("Session ended automatically after {0} minutes of inactivity.", timeoutMinutes), userId:=0, actionType:="SESSION_TIMEOUT")
            AuthenticateUser("")
            lblStatusMessage.Text = String.Format("Session timed out after {0} minutes of inactivity. Tap your RFID card to sign in again.", timeoutMinutes)
        End Sub

        Private Sub ShowOperatorWarningThreadSafe(message As String)
            If Me.IsDisposed OrElse Not Me.IsHandleCreated Then Return
            If Me.InvokeRequired Then
                Me.BeginInvoke(Sub() ShowOperatorWarning(message))
            Else
                ShowOperatorWarning(message)
            End If
        End Sub

        Private Sub ShowOperatorWarning(message As String)
            lblStatusMessage.Text = message
            lblStatusMessage.ForeColor = CivicCalmTheme.ColorWarning
        End Sub

        Private Sub OnCacheSaveFailed(message As String)
            ShowOperatorWarningThreadSafe("WARNING: local offline cache could not be saved (" & message & "). Recent offline work is at risk until this is resolved.")
        End Sub

        ''' <summary>
        ''' The session row is captured at login, while the mirror moves Users rows from
        ''' cache-local ids onto SQL identities. Re-point it at the mirrored row for the same
        ''' badge so UserID stays the SQL identity every write path puts into a foreign key.
        ''' When the mirrored table no longer carries the badge (disabled account) the held
        ''' row is detached and every later read throws RowNotInTableException, so the
        ''' session drops to the logged-out state that every CurrentUser call site checks
        ''' for. The drop is now announced on the status banner: offline enrolments replay
        ''' to SQL before the pull, so a vanished badge means the account was disabled.
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
                lblUserBadge.Text = "[ RFID Logged Out: Access Restricted ]"
                lblUserBadge.ForeColor = CivicCalmTheme.ColorDanger
                lblStatusMessage.Text = "Session ended: your badge is no longer active in the user registry."
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
            If tmrResizeRefresh IsNot Nothing Then
                tmrResizeRefresh.Stop()
                tmrResizeRefresh.Dispose()
                tmrResizeRefresh = Nothing
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
                .Name = "pnlSidebar",
                .Dock = DockStyle.Left,
                .Width = 246,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(12)
            }

            pnlBrand = New Panel With {
                .Dock = DockStyle.Top,
                .Height = 76,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Margin = New Padding(0, 0, 0, 8),
                .Padding = New Padding(2, 4, 2, 4)
            }

            Dim pnlBrandText As New Panel With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(0, 2, 0, 0),
                .BackColor = CivicCalmTheme.ColorSurface
            }
            Dim lblBrand As New Label With {
                .Text = "BTA PARLIAMENT" & vbCrLf & "OFFICE OF THE SG",
                .Font = CivicCalmTheme.FontFormTitle,
                .ForeColor = CivicCalmTheme.ColorPrimary,
                .Dock = DockStyle.Top,
                .Height = 44,
                .TextAlign = ContentAlignment.MiddleLeft
            }
            Dim lblBrandSub As New Label With {
                .Text = "DOCUMENT TRACKING SYSTEM",
                .Font = CivicCalmTheme.FontMicrocopy,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .Dock = DockStyle.Top,
                .Height = 20,
                .TextAlign = ContentAlignment.MiddleLeft
            }
            pnlBrandText.Controls.AddRange(New Control() {lblBrandSub, lblBrand})
            pnlBrand.Controls.Add(pnlBrandText)

            Dim pnlNavStack As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.TopDown,
                .WrapContents = False,
                .AutoScroll = True,
                .BackColor = CivicCalmTheme.ColorSurface
            }

            tipNav = New ToolTip()

            Dim navItems As New List(Of String) From {"Dashboard", "Data Analytics", "Document Registry", "SG Directives", "Search & Storage", "User & RFID Admin", "Audit Trail", "Portal Intake"}
            Dim navIcons As New List(Of String) From {"squares-four", "chart-bar", "file-text", "arrows-split", "magnifying-glass", "user-gear", "shield-check", "globe"}

            For i As Integer = 0 To navItems.Count - 1
                Dim idx As Integer = i
                navNames.Add(navItems(i))
                navIconNames.Add(navIcons(i))
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
                    .BackColor = CivicCalmTheme.ColorSurface,
                    .Image = AppAssets.GetIcon(navIcons(i), 18, CivicCalmTheme.ColorInkMuted),
                    .ImageAlign = ContentAlignment.MiddleLeft,
                    .TextImageRelation = TextImageRelation.ImageBeforeText,
                    .Padding = New Padding(8, 0, 0, 0)
                }
                btn.FlatAppearance.BorderSize = 1
                btn.FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder
                btn.FlatAppearance.MouseOverBackColor = CivicCalmTheme.ColorWell

                ' Feature-gated destinations drop out of the nav when switched off in the
                ' connect dialog's options. Visible toggling keeps every button's index
                ' stable, so the view switcher's fixed cases never shift.
                If navItems(i) = "Data Analytics" AndAlso Not AppSettings.Instance.AnalyticsEnabled Then btn.Visible = False
                If navItems(i) = "Portal Intake" AndAlso Not AppSettings.Instance.PortalSettings.PortalEnabled Then btn.Visible = False

                AddHandler btn.Click, Sub() SwitchNavView(idx)
                navButtons.Add(btn)
                pnlNavStack.Controls.Add(btn)
            Next

            Dim pnlSidebarFooter As New Panel With {
                .Name = "pnlSidebarFooter",
                .Dock = DockStyle.Bottom,
                .Height = 50,
                .Padding = New Padding(0, 6, 0, 0),
                .BackColor = CivicCalmTheme.ColorSurface
            }

            btnCheckUpdate = New Button With {
                .Name = "btnNav_WhatsNew",
                .Text = "  What's New",
                .Dock = DockStyle.Fill,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontBody,
                .TextAlign = ContentAlignment.MiddleLeft,
                .Cursor = Cursors.Hand,
                .UseMnemonic = False,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Image = AppAssets.GetIcon("sparkle", 18, CivicCalmTheme.ColorInkMuted),
                .ImageAlign = ContentAlignment.MiddleLeft,
                .TextImageRelation = TextImageRelation.ImageBeforeText,
                .Padding = New Padding(8, 0, 0, 0)
            }
            btnCheckUpdate.FlatAppearance.BorderSize = 1
            btnCheckUpdate.FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder
            btnCheckUpdate.FlatAppearance.MouseOverBackColor = CivicCalmTheme.ColorWell
            AddHandler btnCheckUpdate.Click, Async Sub()
                btnCheckUpdate.Enabled = False
                Dim origText = btnCheckUpdate.Text
                btnCheckUpdate.Text = "  Checking..."
                Try
                    Await AppUpdateService.CheckAndApplyUpdateAsync(Me, True, Sub()
                        btnCheckUpdate.Text = origText
                        btnCheckUpdate.Enabled = True
                    End Sub)
                Finally
                    btnCheckUpdate.Text = origText
                    btnCheckUpdate.Enabled = True
                End Try
            End Sub
            pnlSidebarFooter.Controls.Add(btnCheckUpdate)

            pnlSidebar.Controls.Add(pnlNavStack)
            pnlSidebar.Controls.Add(pnlSidebarFooter)
            pnlSidebar.Controls.Add(pnlBrand)
            pnlNavStack.BringToFront()
            Me.Controls.Add(pnlSidebar)
        End Sub

        Private Sub SwitchNavView(index As Integer)
            ' Admin and Audit expose every user's record (badge UIDs included) and every
            ' audit event. Evaluated per click so a logout re-gates without rebuilding the
            ' sidebar; the write actions stay gated separately at their own sites.
            Dim navName = navNames(index)
            Dim role = If(CurrentUser IsNot Nothing AndAlso CurrentUser.Table.Columns.Contains("Role") AndAlso Not IsDBNull(CurrentUser("Role")), CurrentUser("Role").ToString(), "")
            Dim privileged As Boolean = (role = "System Administrator" OrElse role = "Secretary-General")
            If (navName = "User & RFID Admin" OrElse navName = "Audit Trail") AndAlso Not privileged Then
                lblStatusMessage.Text = If(CurrentUser Is Nothing,
                    "Sign in with a badge before opening this screen.",
                    "Access Denied: this screen is limited to the System Administrator and the Secretary-General.")
                lblStatusMessage.ForeColor = CivicCalmTheme.ColorDanger
                Return
            End If
            activeNavIndex = index
            For i As Integer = 0 To navButtons.Count - 1
                If i = index Then
                    navButtons(i).BackColor = CivicCalmTheme.ColorPrimarySoft
                    navButtons(i).ForeColor = CivicCalmTheme.ColorPrimary
                    navButtons(i).Font = CivicCalmTheme.FontFieldLabel
                    navButtons(i).FlatAppearance.BorderColor = CivicCalmTheme.ColorPrimary
                    If i < navIconNames.Count Then
                        navButtons(i).Image = AppAssets.GetIcon(navIconNames(i), 18, CivicCalmTheme.ColorPrimary)
                    End If
                Else
                    navButtons(i).BackColor = CivicCalmTheme.ColorSurface
                    navButtons(i).ForeColor = CivicCalmTheme.ColorInkMuted
                    navButtons(i).Font = CivicCalmTheme.FontBody
                    navButtons(i).FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder
                    If i < navIconNames.Count Then
                        navButtons(i).Image = AppAssets.GetIcon(navIconNames(i), 18, CivicCalmTheme.ColorInkMuted)
                    End If
                End If
            Next

            ' Controls.Clear() disposes what it removes, and these view panels are built once, so
            ' remove them without disposing or the next visit re-adds a dead control.
            pnlContent.SuspendLayout()
            Try
                While pnlContent.Controls.Count > 0
                    pnlContent.Controls.RemoveAt(0)
                End While
                Select Case index
                    Case 0 : pnlContent.Controls.Add(viewDashboard)
                    Case 1 : If viewAnalytics IsNot Nothing Then pnlContent.Controls.Add(viewAnalytics)
                    Case 2 : pnlContent.Controls.Add(viewRegistry)
                    Case 3 : pnlContent.Controls.Add(viewDirectives)
                    Case 4 : pnlContent.Controls.Add(viewSearch)
                    Case 5 : pnlContent.Controls.Add(viewAdmin)
                    Case 6 : pnlContent.Controls.Add(viewAudit)
                    Case 7 : If viewPortalIntake IsNot Nothing Then pnlContent.Controls.Add(viewPortalIntake)
                End Select
            Finally
                pnlContent.ResumeLayout(True)
            End Try

            ' The grid balance must measure after layout resumes: inside the suspended pass it
            ' read pre-layout bounds and cached wrong column widths until the next visit.
            RefreshActiveTabGrid()
        End Sub

        Private Sub SetupHeader()
            pnlHeader = New Panel With {
                .Dock = DockStyle.Top,
                .Height = 86,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(20, 8, 20, 8)
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
            tblHeader.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

            Dim pnlHeaderBrand As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 2,
                .RowCount = 1,
                .BackColor = CivicCalmTheme.ColorSurface
            }
            pnlHeaderBrand.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 74.0F))
            pnlHeaderBrand.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            pnlHeaderBrand.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

            Dim picHeaderLogo = AppAssets.CreateLogoPictureBox(64)
            picHeaderLogo.Anchor = AnchorStyles.Left
            picHeaderLogo.Margin = New Padding(0, 0, 10, 0)

            Dim pnlTitles As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.TopDown,
                .WrapContents = False,
                .Margin = New Padding(0),
                .Padding = New Padding(0, 8, 0, 0),
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
                .AutoSize = True,
                .Margin = New Padding(0)
            }

            pnlTitles.Controls.AddRange(New Control() {lblTitle, lblUserBadge})
            pnlHeaderBrand.Controls.Add(picHeaderLogo, 0, 0)
            pnlHeaderBrand.Controls.Add(pnlTitles, 1, 0)

            Dim pnlAuthActions As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.RightToLeft,
                .WrapContents = False,
                .BackColor = CivicCalmTheme.ColorSurface
            }

            btnLogout = New Button With {
                .Text = " &Logout",
                .Size = New Size(105, 36),
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontBody,
                .BackColor = CivicCalmTheme.ColorWell,
                .ForeColor = CivicCalmTheme.ColorInk,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(8, 14, 0, 0),
                .Image = AppAssets.GetIcon("sign-out", 16, CivicCalmTheme.ColorInk),
                .ImageAlign = ContentAlignment.MiddleLeft,
                .TextAlign = ContentAlignment.MiddleCenter,
                .TextImageRelation = TextImageRelation.ImageBeforeText,
                .Padding = New Padding(6, 0, 6, 0)
            }
            btnLogout.FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder
            AddHandler btnLogout.Click, Sub() AuthenticateUser("")

            btnScanRFID = New Button With {
                .Text = " &Tap RFID Smart Card",
                .Size = New Size(205, 36),
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .BackColor = CivicCalmTheme.ColorPrimary,
                .ForeColor = Color.White,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(0, 14, 0, 0),
                .Image = AppAssets.GetIcon("identification-card", 18, Color.White),
                .ImageAlign = ContentAlignment.MiddleLeft,
                .TextAlign = ContentAlignment.MiddleCenter,
                .TextImageRelation = TextImageRelation.ImageBeforeText,
                .Padding = New Padding(8, 0, 8, 0)
            }
            btnScanRFID.FlatAppearance.BorderSize = 0
            AddHandler btnScanRFID.Click, Sub() ShowRFIDLoginDialog()

            pnlAuthActions.Controls.AddRange(New Control() {btnLogout, btnScanRFID})

            tblHeader.Controls.Add(pnlHeaderBrand, 0, 0)
            tblHeader.Controls.Add(pnlAuthActions, 1, 0)
            pnlHeader.Controls.Add(tblHeader)
            tblHeader.BringToFront()
            Me.Controls.Add(pnlHeader)
        End Sub

        Private Sub ShowRFIDLoginDialog()
            Using dlg As New FormLogin()
                If dlg.ShowDialog(Me) = DialogResult.OK AndAlso Not String.IsNullOrEmpty(dlg.ScannedUID) Then
                    AuthenticateUser(dlg.ScannedUID, dlg.PickedOnScreen)
                End If
            End Using
        End Sub

        Public Sub AuthenticateUser(uid As String, Optional simulated As Boolean = False)
            If String.IsNullOrEmpty(uid) Then
                If CurrentUser IsNot Nothing Then
                    Dim who As String = ""
                    Try
                        who = CurrentUser("FullName").ToString()
                    Catch
                    End Try
                    EmbeddedDB.LogAudit(who, "User logged out.", actionType:="LOGOUT")
                End If
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
                    ' The mirror carries IsLocked: a locked account must not unlock itself by
                    ' tapping, and the success path below clears the SQL-side counter, which
                    ' would reset the lock along with it. Refuse here instead.
                    If user.Table.Columns.Contains("IsLocked") AndAlso Not IsDBNull(user("IsLocked")) AndAlso Convert.ToBoolean(user("IsLocked")) Then
                        EmbeddedDB.LogAudit(user("FullName").ToString(), "RFID Badge Tap Blocked: account is locked [Card: " & maskedUid & "]", actionType:="LOGIN_LOCKED")
                        lblStatusMessage.Text = "Access Denied: this account is locked after repeated failed badge reads. Ask a System Administrator to unlock it."
                        lblStatusMessage.ForeColor = CivicCalmTheme.ColorDanger
                        Return
                    End If
                    CurrentUser = user
                    Dim role = user("Role").ToString()
                    Dim office = If(user.Table.Columns.Contains("Office") AndAlso Not IsDBNull(user("Office")), user("Office").ToString(), "")
                    Dim effectiveDesk = If(Not String.IsNullOrWhiteSpace(office), office, role)
                    Dim isGlobal As Boolean = (role = "Secretary-General" OrElse role = "System Administrator" OrElse role = "OSG Chief")
                    Dim displayRole = If(Not String.IsNullOrWhiteSpace(office) AndAlso Not office.Equals(role, StringComparison.OrdinalIgnoreCase), String.Format("{0} - {1}", role.ToUpperInvariant(), office.ToUpperInvariant()), role.ToUpperInvariant())
                    lblUserBadge.Text = String.Format("AUTHENTICATED: {0} [{1}] : {2}", user("FullName").ToString().ToUpperInvariant(), displayRole, If(isGlobal, "GLOBAL ACCESS", "SECTION DESK VIEW"))
                    lblUserBadge.ForeColor = CivicCalmTheme.ColorPrimary
                    lblStatusMessage.Text = "Authenticated: " & user("FullName").ToString()
                    lblStatusMessage.ForeColor = CivicCalmTheme.ColorInk
                    EmbeddedDB.LogAudit(user("FullName").ToString(), If(simulated,
                        "On-screen badge selection authenticated [Card: " & maskedUid & "] (no reader tap)",
                        "RFID Badge Tap Authenticated [Card: " & maskedUid & "]"), actionType:="LOGIN_SUCCESS")

                    ' Successful tap clears the SQL-side failed tap counter so the per-user
                    ' lockout (tbl_Users.FailedTapCount) tracks live state, not stale counts.
                    If AppStartup.UserRepo IsNot Nothing AndAlso Program.IsDatabaseConnected Then
                        Try
                            Dim userId = If(user.Table.Columns.Contains("UserID") AndAlso Not IsDBNull(user("UserID")), Convert.ToInt32(user("UserID")), 0)
                            If userId > 0 Then AppStartup.UserRepo.ResetFailedTaps(userId)
                        Catch
                        End Try
                    End If

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
                    ' A cache miss with a live SQL host means one of two things: a truly
                    ' unknown card, or a badge the server knows but this seat's Users mirror
                    ' predates. GetByCardPublicID only matches active badges, so a hit here is
                    ' the stale-mirror case: the owner must not be punished for the seat's
                    ' lag. Refresh the mirror and invite a second tap instead of counting a
                    ' failed tap against the account.
                    Dim staleMirrorResolved As Boolean = False
                    If AppStartup.UserRepo IsNot Nothing AndAlso Program.IsDatabaseConnected Then
                        Try
                            Dim knownUser = AppStartup.UserRepo.GetByCardPublicID(uid.Trim().ToUpperInvariant())
                            If knownUser IsNot Nothing Then
                                If Program.Coordinator IsNot Nothing Then Program.Coordinator.RefreshFromServer("Users")
                                EmbeddedDB.LogAudit("SYSTEM", "Badge recognized on the server; this seat's staff mirror was stale and has been refreshed [Card: " & maskedUid & "]", actionType:="MIRROR_REFRESH")
                                lblStatusMessage.Text = "This badge is valid on the server; the seat's staff list was stale and has been refreshed. Tap the badge again."
                                lblStatusMessage.ForeColor = CivicCalmTheme.ColorWarning
                                staleMirrorResolved = True
                            End If
                        Catch
                        End Try
                    End If
                    If Not staleMirrorResolved Then
                        If EmbeddedDB.IsTerminalLockedOut() Then
                            lblStatusMessage.Text = "Security Lockout: Terminal locked for 5 minutes due to 5 consecutive failed card reads."
                            MessageBox.Show("Terminal has been temporarily locked out due to 5 consecutive failed RFID smart card badge reads. Please notify the System Administrator or wait 5 minutes.", "Security Lockout", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Else
                            lblStatusMessage.Text = "Access Denied: Unrecognized RFID card [" & maskedUid & "]"
                            MessageBox.Show("Unrecognized RFID Smart Card Badge UID: " & maskedUid, "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        End If
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

        ''' <summary>
        ''' Portal status strings arrive from background tasks; marshal the label write onto
        ''' the UI thread. The Nothing guard covers callbacks that fire before the intake
        ''' view exists.
        ''' </summary>
        Private Sub SetPortalStatusThreadSafe(message As String)
            If Me.InvokeRequired Then
                Me.Invoke(Sub()
                              If lblPortalStatus IsNot Nothing Then lblPortalStatus.Text = message
                          End Sub)
            Else
                If lblPortalStatus IsNot Nothing Then lblPortalStatus.Text = message
            End If
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
            ' The row can vanish between grid paint and click when a sync deletes or
            ' re-filters it; the dialog constructor would dispose itself and ShowDialog
            ' would throw on a dead object.
            If EmbeddedDB.GetDocumentByID(docId) Is Nothing Then
                lblStatusMessage.Text = "This document is no longer available. Refresh the grid and try again."
                Return
            End If
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
            If Me.WindowState = FormWindowState.Minimized OrElse Me.ClientRectangle.Width <= 0 OrElse Me.ClientRectangle.Height <= 0 Then
                Return
            End If

            Dim selSec As String = If(cmbDashSection IsNot Nothing AndAlso cmbDashSection.SelectedItem IsNot Nothing, cmbDashSection.SelectedItem.ToString(), "")
            Dim selCat As String = If(cmbDashCategory IsNot Nothing AndAlso cmbDashCategory.SelectedItem IsNot Nothing, cmbDashCategory.SelectedItem.ToString(), "")
            ' Reuse one view across sync ticks: the mirror merge updates the underlying table
            ' in place, so a bound view stays current without a rebuild. Recreate only when
            ' the session or the dashboard filters change.
            Dim sessionKey As String = "(out)"
            If CurrentUser IsNot Nothing Then
                Dim office = If(CurrentUser.Table.Columns.Contains("Office") AndAlso Not IsDBNull(CurrentUser("Office")), CurrentUser("Office").ToString(), "")
                sessionKey = CurrentUser("FullName").ToString() & "|" & CurrentUser("Role").ToString() & "|" & office
            End If
            Dim viewKey = sessionKey & "|" & selSec & "|" & selCat
            If _activeDocsView Is Nothing OrElse viewKey <> _activeDocsViewKey Then
                _activeDocsView = EmbeddedDB.GetVisibleDocuments(CurrentUser, selSec, selCat)
                _activeDocsViewKey = viewKey
            End If
            Dim visibleDocs = _activeDocsView

            Select Case activeNavIndex
                Case 0 : RefreshDashboardGrid(visibleDocs)
                Case 1 : RefreshAnalyticsView()
                Case 2 : RefreshRegistryGrid(visibleDocs)
                Case 3 : RefreshDirectivesGrid(visibleDocs)
                Case 4 : RefreshSearchGrid(visibleDocs)
                Case 5 : RefreshAdminGrid()
                Case 6 : RefreshAuditGrid()
                Case 7 : RefreshPortalQueue()
            End Select

            lblStatusClock.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        End Sub

        ''' <summary>
        ''' Binds a grid, applies its column formatter, switches the empty/populated watermark,
        ''' and (when a noun is given) mirrors the row count into the footer count badge.
        ''' Assigning a DataSource makes the grid rebuild every column and row, so the rebind
        ''' is skipped while the view instance is unchanged; the mirror merge keeps a bound
        ''' view current through its own row events. forceBind re-runs the formatter for grids
        ''' whose layout is computed from row values (seat staleness).
        ''' </summary>
        Private Sub BindGridWithState(dgv As DataGridView, watermark As Label, dataSource As Object, formatter As Action(Of DataGridView), countNoun As String, Optional emptyMessage As String = Nothing, Optional forceBind As Boolean = False)
            If forceBind OrElse Not Object.ReferenceEquals(dgv.DataSource, dataSource) Then
                dgv.DataSource = dataSource
                formatter(dgv)
            End If
            If dgv.Rows.Count = 0 Then
                DataGridStyler.SetEmptyState(dgv, watermark, emptyMessage)
            Else
                DataGridStyler.SetPopulatedState(dgv, watermark)
            End If
            If countNoun <> "" Then
                lblStatusCount.Text = dgv.Rows.Count.ToString() & " " & countNoun
            End If
        End Sub

        Private Sub RefreshDashboardGrid(visibleDocs As DataView)
            ' The stat cards are office-wide numbers (the labels say so), while the grid below
            ' honors the section and category filters. The row walk runs under the store lock:
            ' the replay worker mutates these same rows on sync ticks.
            Dim allDocs = EmbeddedDB.DataSet.Tables("Documents")
            Dim forReviewCount As Integer = 0
            Dim forRevisionCount As Integer = 0
            Dim approvedReleasedCount As Integer = 0
            SyncLock EmbeddedDB.SyncRoot
                For Each row As DataRow In allDocs.Rows
                    Dim st = row("CurrentStatus").ToString()
                    If st.IndexOf("REVIEW", StringComparison.OrdinalIgnoreCase) >= 0 Then forReviewCount += 1
                    If st.IndexOf("REVISION", StringComparison.OrdinalIgnoreCase) >= 0 Then forRevisionCount += 1
                    If st.Equals("APPROVED", StringComparison.OrdinalIgnoreCase) OrElse st.Equals("RELEASED", StringComparison.OrdinalIgnoreCase) Then approvedReleasedCount += 1
                Next
            End SyncLock

            lblStatTotalDocs.Text = allDocs.Rows.Count.ToString()
            lblStatDirectives.Text = forReviewCount.ToString()
            lblStatActiveRoute.Text = forRevisionCount.ToString()
            lblStatVaultStorage.Text = approvedReleasedCount.ToString()
            BindGridWithState(dgvDashRecent, lblDashWatermark, visibleDocs, AddressOf DataGridStyler.FormatDocumentColumns, "")
        End Sub

        Private Sub RefreshRegistryGrid(visibleDocs As DataView)
            BindGridWithState(dgvRegistry, lblRegistryWatermark, visibleDocs, AddressOf DataGridStyler.FormatDocumentColumns, "Documents")
            LoadRegistrySuggestions()
        End Sub

        Private Sub RefreshDirectivesGrid(visibleDocs As DataView)
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
            BindGridWithState(dgvDirectives, lblDirectivesWatermark, dtDirectives, AddressOf DataGridStyler.FormatDirectiveColumns, "Directives", "No active action directives found. Select a document above and click 'Log Action Directive' to issue an executive directive.")
        End Sub

        Private Sub RefreshSearchGrid(visibleDocs As DataView)
            ' While a search query is active the grid is owned by OnSearch; a sync tick
            ' must not reset what the operator is looking at. Same contract as RefreshAuditGrid.
            If txtSearchKey IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(txtSearchKey.Text) Then Return
            BindGridWithState(dgvSearch, lblSearchWatermark, visibleDocs, AddressOf DataGridStyler.FormatDocumentColumns, "Documents")
        End Sub

        Private Sub RefreshAdminGrid()
            Dim dtUsers = EmbeddedDB.DataSet.Tables("Users")
            BindGridWithState(dgvUsers, lblUsersWatermark, dtUsers, AddressOf DataGridStyler.FormatUserColumns, "Users", "No users registered.")
            If dgvSeats IsNot Nothing Then
                Dim dtHeartbeat = EmbeddedDB.DataSet.Tables("Heartbeat")
                BindGridWithState(dgvSeats, lblSeatsWatermark, dtHeartbeat, AddressOf DataGridStyler.FormatSeatColumns, "Seats",
                                  "No workstation has reported a sync yet. Seats appear here after their first sync; Stale means the seat has been silent past three intervals (offline by design).",
                                  forceBind:=True)
            End If
        End Sub

        Private Sub RefreshAuditGrid()
            ' While a filter query is active the grid is owned by OnFilterAudit; a sync tick
            ' must not reset what the operator is looking at.
            If txtAuditSearch Is Nothing OrElse String.IsNullOrWhiteSpace(txtAuditSearch.Text) Then
                ' One sorted view for the life of the form: the DataView maintains its sort
                ' index incrementally as the mirror appends rows, so new events surface at
                ' the top without rebuilding an index over the whole trail on every tick.
                If _auditView Is Nothing Then
                    _auditView = New DataView(EmbeddedDB.DataSet.Tables("AuditTrail")) With {.Sort = "AuditID DESC"}
                End If
                BindGridWithState(dgvAudit, lblAuditWatermark, _auditView, AddressOf DataGridStyler.FormatAuditColumns, "Audit Events", "No audit trail events recorded.")
            End If
        End Sub

        Protected Overrides Sub OnResize(e As EventArgs)
            MyBase.OnResize(e)
            If Me.WindowState <> FormWindowState.Minimized AndAlso Me.IsHandleCreated AndAlso Not Me.IsDisposed Then
                ScheduleResizeRefresh()
            End If
        End Sub

        ' Resize raises one message per pixel of a drag; refreshing per message made large
        ' grids stutter. One deferred refresh after the drag settles is visually identical.
        Private Sub ScheduleResizeRefresh()
            If tmrResizeRefresh Is Nothing Then
                tmrResizeRefresh = New Timer With {.Interval = 200}
                AddHandler tmrResizeRefresh.Tick,
                    Sub()
                        tmrResizeRefresh.Stop()
                        If Not Me.IsDisposed AndAlso Me.IsHandleCreated Then RefreshActiveTabGrid()
                    End Sub
            End If
            tmrResizeRefresh.Stop()
            tmrResizeRefresh.Start()
        End Sub

        Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
            If keyData = (Keys.Alt Or Keys.C) Then
                ClearActiveViewFilters()
                Return True
            End If
            If keyData = (Keys.Alt Or Keys.R) AndAlso activeViewRetry IsNot Nothing Then
                ' Only claim Alt+R while a view actually has a retry to run: command keys
                ' resolve before mnemonics, so an unconditional claim deadens every &R
                ' control on every view (the Registry's own Register button among them).
                TryRetryActiveView()
                Return True
            End If
            Return MyBase.ProcessCmdKey(msg, keyData)
        End Function

        Private Sub ClearActiveViewFilters()
            Select Case activeNavIndex
                Case 0
                    btnDashResetFilters.PerformClick()
                Case 1
                    If cmbAnalyticsTimeframe IsNot Nothing Then cmbAnalyticsTimeframe.SelectedIndex = 2
                    If cmbAnalyticsSection IsNot Nothing Then cmbAnalyticsSection.SelectedIndex = 0
                    RefreshAnalyticsView()
                Case 2
                    btnDashResetFilters.PerformClick()
                Case 4
                    btnClearSearch.PerformClick()
                Case 6
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
                If tmrResizeRefresh IsNot Nothing Then
                    tmrResizeRefresh.Stop()
                    tmrResizeRefresh.Dispose()
                    tmrResizeRefresh = Nothing
                End If
            End If
            MyBase.Dispose(disposing)
        End Sub
    End Class
End Namespace
