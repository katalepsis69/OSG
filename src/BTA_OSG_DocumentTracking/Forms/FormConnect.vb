Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Microsoft.Data.SqlClient

Namespace BTA_OSG
    ''' <summary>
    ''' The one setup surface, connect-first: a seat types the shared password and the
    ''' scan finds the server. "Set up this server" exists only where a local SQL Server
    ''' answered but rejected the login, because provisioning can only run on the
    ''' machine that hosts the instance.
    ''' </summary>
    Public Class FormConnect
        Inherits Form

        ' The login every workgroup seat shares; the password typed here is the only
        ' secret an operator ever needs. SQL-auth only: workgroup seats cannot carry
        ' Windows credentials to the server.
        Friend Const AppLoginName As String = "bta_app"
        Friend Const AppDatabaseName As String = "BTA_OSG_DB"

        ' The published demo key is refused so a repo reader cannot reach the intake
        ' queue; an empty one would leave the portal API open too.
        Private Const PublishedDemoBridgeKey As String = "4e5a9b71f92e482db591c890a82b9a714e5a9b71f92e482db591c890a82b9a71"

        ' The built-in portal server answers on one fixed local port; the desktop never
        ' points anywhere else, so the URL is a constant, not a field.
        Friend Const LocalPortalBaseUrl As String = "http://localhost:8085"

        Private _busy As Boolean = False
        Private _advancedVisible As Boolean = False
        Private _probeServer As String = ""
        Private _justPrepared As Boolean = False
        Private _switchingFromDemo As Boolean = False
        Private _startedPortal As Boolean = False
        Private _startedAnalytics As Boolean = True

        Private lblStatus As Label
        Private txtPassword As TextBox
        Private btnConnect As Button
        Private btnPrepare As Button
        Private lnkOptions As LinkLabel
        Private lnkDemo As LinkLabel
        Private pnlAdvanced As Panel
        Private txtServer As TextBox
        Private txtPort As TextBox
        Private chkPortal As CheckBox
        Private chkAnalytics As CheckBox
        Private btnClose As Button

        Public Sub New()
            InitializeForm()
            LoadCurrentSettings()
        End Sub

        ''' <summary>
        ''' FormConnect rewrites the SQL connection string and the portal switch, so once a
        ''' workstation is configured it must not open without a System Administrator badge.
        ''' First-run (unconfigured) stays open so a fresh install can be set up at all.
        ''' </summary>
        Public Shared Function AuthorizeConfiguredChange() As Boolean
            If Not AppSettings.Instance.IsConfigured Then Return True
            Using dlg As New FormLogin()
                If dlg.ShowDialog() <> DialogResult.OK OrElse String.IsNullOrWhiteSpace(dlg.ScannedUID) Then Return False
                Dim user = EmbeddedDB.AuthenticateRFID(dlg.ScannedUID)
                If user Is Nothing Then
                    MessageBox.Show("Unrecognized card. Station setup requires a System Administrator badge.", "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return False
                End If
                Dim role = user("Role").ToString()
                If role <> "System Administrator" AndAlso role <> "SYSADMIN" Then
                    EmbeddedDB.LogAudit(user("FullName").ToString(), "Denied station setup change (role: " & role & ")", actionType:="SETUP_DENIED")
                    MessageBox.Show("Only a System Administrator may change station configuration.", "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return False
                End If
                EmbeddedDB.LogAudit(user("FullName").ToString(), "Authorized station setup change", actionType:="SETUP_AUTHORIZED")
                Return True
            End Using
        End Function

        Private Sub InitializeForm()
            AppAssets.ApplyFormIcon(Me)
            Me.Text = "BTA OSG Document Tracking : Connect"
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.BackColor = CivicCalmTheme.ColorCanvas
            Me.Font = CivicCalmTheme.FontBody
            ' Fixed client size switched between two measured heights: AutoSize on a form
            ' with nested percent-column tables under-reports the preferred width and
            ' clipped the right column off the password, port, and bridge rows, so the
            ' wizard's proven 620 is pinned and the options panel just swaps the height.
            Me.ClientSize = New Size(620, 350)

            Dim pnlRoot As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .Padding = New Padding(20)
            }
            For i As Integer = 0 To 5
                pnlRoot.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            Next
            ' The filler row absorbs the slack between the content and the bottom-docked
            ' footer, so neither height can clip a row.
            pnlRoot.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

            ' Section 1: Header
            Dim pnlHeader As New TableLayoutPanel With {
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .ColumnCount = 2
            }
            pnlHeader.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 58.0F))
            pnlHeader.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            Dim picLogo = AppAssets.CreateLogoPictureBox(48)
            picLogo.Anchor = AnchorStyles.Left
            picLogo.Margin = New Padding(0, 0, 10, 0)
            Dim pnlHeaderText As New Panel With {
                .Dock = DockStyle.Fill,
                .AutoSize = True,
                .Padding = New Padding(0, 4, 0, 0)
            }
            Dim lblTitle As New Label With {
                .Text = "Connect to the OSG Database",
                .Font = CivicCalmTheme.FontFormTitle,
                .ForeColor = CivicCalmTheme.ColorPrimary,
                .Dock = DockStyle.Top,
                .AutoSize = True
            }
            Dim lblSubtitle As New Label With {
                .Text = "One shared password. The scan finds the office server for you.",
                .Font = CivicCalmTheme.FontFieldLabel,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .Dock = DockStyle.Top,
                .AutoSize = True
            }
            pnlHeaderText.Controls.AddRange(New Control() {lblSubtitle, lblTitle})
            pnlHeader.Controls.Add(picLogo, 0, 0)
            pnlHeader.Controls.Add(pnlHeaderText, 1, 0)

            ' Section 2: Status line. Every state message names the cause and the next
            ' action; the minimum height reserves two wrapped lines so the dialog does
            ' not jump when a longer message arrives.
            lblStatus = New Label With {
                .Text = "",
                .Font = CivicCalmTheme.FontBody,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .AutoSize = False,
                .Dock = DockStyle.Fill,
                .Height = 44,
                .Margin = New Padding(0, 10, 0, 0)
            }

            ' Section 3: The password is the only always-visible input. Username and
            ' database are constants of the product, not decisions for the operator.
            Dim pnlPassword As New TableLayoutPanel With {
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .ColumnCount = 2,
                .Padding = New Padding(0, 8, 0, 0)
            }
            pnlPassword.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 32.0F))
            pnlPassword.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 68.0F))
            Dim lblPassword As New Label With {
                .Text = "Shared password:",
                .Font = CivicCalmTheme.FontFieldLabel,
                .ForeColor = CivicCalmTheme.ColorInk,
                .Dock = DockStyle.Fill,
                .TextAlign = ContentAlignment.MiddleLeft,
                .Margin = New Padding(0, 2, 8, 2)
            }
            txtPassword = New TextBox With {
                .Dock = DockStyle.Fill,
                .UseSystemPasswordChar = True,
                .Margin = New Padding(0, 2, 0, 2),
                .BackColor = CivicCalmTheme.ColorSurface,
                .ForeColor = CivicCalmTheme.ColorInk,
                .BorderStyle = BorderStyle.FixedSingle,
                .TabIndex = 0
            }
            pnlPassword.Controls.Add(lblPassword, 0, 0)
            pnlPassword.Controls.Add(txtPassword, 1, 0)

            ' Section 4: Actions. Connect both tests and saves; there is no second step.
            Dim pnlActions As New FlowLayoutPanel With {
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .FlowDirection = FlowDirection.LeftToRight,
                .WrapContents = False,
                .Padding = New Padding(0, 10, 0, 0)
            }
            btnConnect = New Button With {
                .Text = "&Connect",
                .Size = New Size(150, 34),
                .BackColor = CivicCalmTheme.ColorPrimary,
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .TabIndex = 1
            }
            btnConnect.FlatAppearance.BorderSize = 0
            AddHandler btnConnect.Click, AddressOf OnConnect

            ' Hidden until a probe fails against this machine's own SQL Server: on a
            ' seat the button never appears, so nobody is asked to administer a server.
            btnPrepare = New Button With {
                .Text = "Set &up this server",
                .Size = New Size(180, 34),
                .Visible = False,
                .BackColor = CivicCalmTheme.ColorWell,
                .ForeColor = CivicCalmTheme.ColorInk,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(10, 0, 0, 0),
                .TabIndex = 2
            }
            btnPrepare.FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder
            AddHandler btnPrepare.Click, AddressOf OnPrepareServer
            pnlActions.Controls.Add(btnConnect)
            pnlActions.Controls.Add(btnPrepare)

            ' Section 5: The two escapes. Options covers a scan that cannot find the
            ' server; demo covers an office with no server yet at all.
            Dim pnlLinks As New FlowLayoutPanel With {
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .FlowDirection = FlowDirection.LeftToRight,
                .WrapContents = False,
                .Padding = New Padding(0, 8, 0, 0)
            }
            lnkOptions = New LinkLabel With {
                .Text = "Server address and other settings",
                .AutoSize = True,
                .LinkColor = CivicCalmTheme.ColorInfo,
                .ActiveLinkColor = CivicCalmTheme.ColorPrimary,
                .VisitedLinkColor = CivicCalmTheme.ColorInfo,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(0, 3, 18, 0),
                .TabIndex = 3
            }
            AddHandler lnkOptions.LinkClicked, Sub() ToggleAdvanced()
            lnkDemo = New LinkLabel With {
                .Text = "Run without a server (demo mode)",
                .AutoSize = True,
                .LinkColor = CivicCalmTheme.ColorInfo,
                .ActiveLinkColor = CivicCalmTheme.ColorPrimary,
                .VisitedLinkColor = CivicCalmTheme.ColorInfo,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(0, 3, 0, 0),
                .TabIndex = 4
            }
            AddHandler lnkDemo.LinkClicked, AddressOf OnDemoLink
            pnlLinks.Controls.Add(lnkOptions)
            pnlLinks.Controls.Add(lnkDemo)

            ' Section 6: Collapsed options. The rare paths live here: a typed address,
            ' the RFID reader kind, and the thesis-only portal intake.
            pnlAdvanced = New Panel With {
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .Visible = False,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(14),
                .Margin = New Padding(0, 12, 0, 0)
            }
            BuildAdvancedFields()

            ' Section 7: Footer, docked to the bottom of the form rather than living in
            ' the content table, so it stays put at both heights.
            Dim pnlFooter As New FlowLayoutPanel With {
                .Dock = DockStyle.Bottom,
                .AutoSize = True,
                .FlowDirection = FlowDirection.RightToLeft,
                .WrapContents = False,
                .Padding = New Padding(0, 12, 0, 6)
            }
            btnClose = New Button With {
                .Text = "&Close",
                .Size = New Size(100, 32),
                .BackColor = CivicCalmTheme.ColorWell,
                .ForeColor = CivicCalmTheme.ColorInk,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .TabIndex = 30
            }
            btnClose.FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder
            AddHandler btnClose.Click, Sub()
                                           Me.DialogResult = DialogResult.Cancel
                                           Me.Close()
                                       End Sub
            pnlFooter.Controls.Add(btnClose)

            pnlRoot.Controls.Add(pnlHeader, 0, 0)
            pnlRoot.Controls.Add(lblStatus, 0, 1)
            pnlRoot.Controls.Add(pnlPassword, 0, 2)
            pnlRoot.Controls.Add(pnlActions, 0, 3)
            pnlRoot.Controls.Add(pnlLinks, 0, 4)
            pnlRoot.Controls.Add(pnlAdvanced, 0, 5)

            ' Fill added before the bottom-docked footer: the docking engine processes
            ' the last control first, so the footer claims its strip and the root fills
            ' the rest.
            Me.Controls.Add(pnlRoot)
            Me.Controls.Add(pnlFooter)
            Me.AcceptButton = btnConnect
            Me.CancelButton = btnClose
            Me.ActiveControl = txtPassword
        End Sub

        Private Sub BuildAdvancedFields()
            ' Field rows live in a percent-column table rather than at fixed points: at
            ' 125% or 150% display scaling the AutoSize labels grow, and fixed
            ' x-coordinates would let them collide with the inputs (DESIGN.md 2.2).
            Dim tbl As New TableLayoutPanel With {
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .BackColor = CivicCalmTheme.ColorSurface,
                .ColumnCount = 4
            }
            tbl.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 26.0F))
            tbl.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 42.0F))
            tbl.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 12.0F))
            tbl.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 20.0F))

            Dim FieldLabel = Function(text As String, font As Font, color As Color) As Label
                                 Return New Label With {
                                     .Text = text,
                                     .Font = font,
                                     .ForeColor = color,
                                     .Dock = DockStyle.Fill,
                                     .TextAlign = ContentAlignment.MiddleLeft,
                                     .Margin = New Padding(0, 2, 8, 2)
                                 }
                             End Function

            Dim NewInput = Function(text As String, tabIndex As Integer) As TextBox
                               Return New TextBox With {
                                   .Text = text,
                                   .TabIndex = tabIndex,
                                   .Dock = DockStyle.Fill,
                                   .Margin = New Padding(0, 2, 8, 2),
                                   .BackColor = CivicCalmTheme.ColorSurface,
                                   .ForeColor = CivicCalmTheme.ColorInk,
                                   .BorderStyle = BorderStyle.FixedSingle
                               }
                           End Function

            Dim r As Integer = 0

            Dim lblAdvancedTitle As New Label With {
                .Text = "SERVER ADDRESS && OTHER SETTINGS",
                .Font = CivicCalmTheme.FontSectionHeader,
                .ForeColor = CivicCalmTheme.ColorPrimary,
                .AutoSize = True,
                .Margin = New Padding(0, 2, 0, 8)
            }
            tbl.Controls.Add(lblAdvancedTitle, 0, r)
            tbl.SetColumnSpan(lblAdvancedTitle, 4)
            r += 1

            txtServer = NewInput("", 10)
            txtPort = NewInput("1433", 11)
            tbl.Controls.Add(FieldLabel("Server Address:", CivicCalmTheme.FontFieldLabel, CivicCalmTheme.ColorInk), 0, r)
            tbl.Controls.Add(txtServer, 1, r)
            tbl.Controls.Add(FieldLabel("Port:", CivicCalmTheme.FontFieldLabel, CivicCalmTheme.ColorInk), 2, r)
            tbl.Controls.Add(txtPort, 3, r)
            r += 1

            ' One optional-features row: the portal checkbox is the whole portal surface
            ' (the local server always answers on the same port and the bridge key
            ' generates itself), and analytics hides its own nav entry when unchecked.
            Dim pnlFeatures As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .Margin = New Padding(0, 6, 0, 4),
                .BackColor = Color.Transparent
            }
            chkPortal = New CheckBox With {
                .Text = "Enable Citizen Web Portal Intake",
                .Font = CivicCalmTheme.FontFieldLabel,
                .ForeColor = CivicCalmTheme.ColorInk,
                .AutoSize = True,
                .Margin = New Padding(0, 4, 18, 0),
                .TabIndex = 14
            }
            chkAnalytics = New CheckBox With {
                .Text = "Show Data Analytics",
                .Font = CivicCalmTheme.FontFieldLabel,
                .ForeColor = CivicCalmTheme.ColorInk,
                .AutoSize = True,
                .Checked = True,
                .Margin = New Padding(0, 4, 0, 0),
                .TabIndex = 15
            }
            pnlFeatures.Controls.Add(chkPortal)
            pnlFeatures.Controls.Add(chkAnalytics)
            tbl.Controls.Add(pnlFeatures, 0, r)
            tbl.SetColumnSpan(pnlFeatures, 4)

            pnlAdvanced.Controls.Add(tbl)
        End Sub

        Private Sub LoadCurrentSettings()
            Dim s = AppSettings.Instance
            _startedPortal = s.PortalSettings.PortalEnabled
            _startedAnalytics = s.AnalyticsEnabled

            ' The saved address is offered whether or not the station is currently using it:
            ' switching to demo mode keeps the connection string, and making an operator retype
            ' a server the station already knew is how a demo-to-office switch stalls.
            If s.IsConfigured Then
                Try
                    Dim b As New SqlConnectionStringBuilder(s.DatabaseSettings.ConnectionString)
                    Dim srv = b.DataSource
                    If srv.StartsWith("tcp:", StringComparison.OrdinalIgnoreCase) Then srv = srv.Substring(4)
                    Dim srvParts = srv.Split(","c)
                    txtServer.Text = srvParts(0)
                    ' A saved named instance carries no port and must not inherit the 1433
                    ' suggestion, which BuildDataSource would then append to it.
                    If srvParts.Length > 1 Then
                        txtPort.Text = srvParts(1)
                    ElseIf srvParts(0).Contains("\"c) Then
                        txtPort.Text = ""
                    Else
                        txtPort.Text = "1433"
                    End If
                Catch parseEx As Exception
                    ' A malformed saved string must not pose as the saved configuration:
                    ' the operator could save the defaults over the real settings without
                    ' noticing. Show the options panel so the fields are visibly empty.
                    txtServer.Text = ""
                    txtPort.Text = ""
                    SetStatus("The saved connection string could not be read (" & parseEx.GetType().Name & "). Enter the server address below.", CivicCalmTheme.ColorWarning)
                    ToggleAdvanced()
                End Try
            Else
                txtServer.Text = ""
                txtPort.Text = "1433"
            End If
            chkPortal.Checked = s.PortalSettings.PortalEnabled
            chkAnalytics.Checked = s.AnalyticsEnabled

            If Not s.IsConfigured Then
                SetStatus("Type the shared office password, then press Connect. The scan searches the network and remembers the server it finds.", CivicCalmTheme.ColorInkMuted)
            ElseIf s.DatabaseSettings.UseSqlServer Then
                SetStatus("This station was connected to " & If(txtServer.Text.Length > 0, txtServer.Text, "a saved server") & ". Leave the password blank to keep the saved sign-in.", CivicCalmTheme.ColorInkMuted)
            Else
                _switchingFromDemo = True
                SetStatus("This station is running in demo mode. Type the shared office password and press Connect to switch it to the office server.", CivicCalmTheme.ColorInkMuted)
            End If
        End Sub

        Private Sub ToggleAdvanced()
            _advancedVisible = Not _advancedVisible
            pnlAdvanced.Visible = _advancedVisible
            lnkOptions.Text = If(_advancedVisible, "Hide server address and other settings", "Server address and other settings")
            ' Two measured heights: the collapsed dialog stays compact and the options
            ' panel gets its own room instead of AutoSize guesswork.
            Me.ClientSize = New Size(620, If(_advancedVisible, 520, 350))
        End Sub

        ''' <summary>
        ''' The data source for a typed host: a port appends as host,port, while a named
        ''' instance resolves its own port through the SQL Browser service and must never
        ''' get one appended.
        ''' </summary>
        Friend Shared Function BuildDataSource(server As String, port As String) As String
            Dim host = server.Trim()
            Dim portText = If(port, "").Trim()
            If portText.Length = 0 OrElse host.Contains("\"c) Then Return host
            Return host & "," & portText
        End Function

        Friend Shared Function BuildConnectionString(server As String, port As String, user As String, password As String) As String
            Dim b As New SqlConnectionStringBuilder() With {
                .DataSource = BuildDataSource(server, port),
                .InitialCatalog = AppDatabaseName,
                .Encrypt = True,
                .TrustServerCertificate = True,
                .ConnectTimeout = 4,
                .MaxPoolSize = 100,
                .IntegratedSecurity = False,
                .UserID = user,
                .Password = password
            }
            Return b.ConnectionString
        End Function

        ''' <summary>
        ''' The same host, but Windows-authenticated against master: on the server
        ''' console of a fresh Express install the current user is sysadmin. Local
        ''' connections answer over shared memory, so this works even before the TCP
        ''' settings and the firewall are in place.
        ''' </summary>
        Friend Shared Function BuildAdminConnectionString(server As String, port As String) As String
            Dim b As New SqlConnectionStringBuilder() With {
                .DataSource = BuildDataSource(server, port),
                .InitialCatalog = "master",
                .Encrypt = True,
                .TrustServerCertificate = True,
                .ConnectTimeout = 15,
                .IntegratedSecurity = True
            }
            Return b.ConnectionString
        End Function

        Private Async Sub OnConnect(sender As Object, e As EventArgs)
            If _busy Then Return
            Await RunConnectAsync()
        End Sub

        Private Async Function RunConnectAsync() As Task
            Dim credentials = EffectiveCredentials()
            If credentials.Password.Length = 0 Then
                SetStatus("Type the shared office password to connect.", CivicCalmTheme.ColorWarning)
                txtPassword.Focus()
                Return
            End If

            SetBusy(True)
            Dim typedServer = txtServer.Text.Trim()
            SetStatus("Searching the network for the OSG server...", CivicCalmTheme.ColorInfo)

            Dim result As ConnectResult
            Try
                result = Await Task.Run(Function() ProbeTargets(typedServer, txtPort.Text, credentials.User, credentials.Password))
            Catch ex As Exception
                SetStatus("The search did not finish (" & ex.GetType().Name & "). Enter the server address below and try again.", CivicCalmTheme.ColorDanger)
                Return
            Finally
                ' Without this the dialog keeps every control but Close disabled and reads as
                ' frozen, which is what an operator meets if the scan throws.
                SetBusy(False)
            End Try

            Select Case result.Outcome
                Case ConnectOutcome.Connected
                    Dim others = If(result.OtherServers.Count > 0,
                                    " Also found a database on: " & String.Join(", ", result.OtherServers.ToArray()) & ".",
                                    "")
                    SetStatus("Connected. Found the OSG database on " & result.Server & "." & others, CivicCalmTheme.ColorPrimary)
                    SaveSqlServerSettings(result.Server, result.Port, credentials.User, credentials.Password)
                Case ConnectOutcome.SchemaIncomplete
                    OfferPrepare(result, "The server at " & result.Server & " answered, but its database is missing parts of the schema.")
                Case ConnectOutcome.Rejected
                    OfferPrepare(result, "The server at " & result.Server & " rejected this password. On a seat, confirm it with whoever set up the server.")
                Case ConnectOutcome.MissingDatabase
                    OfferPrepare(result, "The server at " & result.Server & " answered but has no OSG database yet.")
                Case Else
                    ' No host answered, so the password was never tested. Saying so is the
                    ' difference between a stalled station and a wrong credential.
                    Dim notChecked = " so this password could not be checked, and this station keeps the mode it already had."
                    If _advancedVisible Then
                        SetStatus("No server answered at " & If(typedServer.Length > 0, typedServer, "the typed address") & ":" & txtPort.Text.Trim() & "," & notChecked & " Check the address and port.", CivicCalmTheme.ColorWarning)
                    Else
                        SetStatus("Nothing on the network is running the office server," & notChecked & " Enter the server address below." & If(_switchingFromDemo, " To stay in demo mode, close this window.", ""), CivicCalmTheme.ColorWarning)
                        ToggleAdvanced()
                    End If
            End Select
        End Function

        Private Enum ConnectOutcome
            Connected
            SchemaIncomplete
            Rejected
            MissingDatabase
            Unreachable
        End Enum

        Private Class ConnectResult
            Public Property Outcome As ConnectOutcome
            Public Property Server As String = ""
            Public Property Port As Integer? = Nothing
            Public Property OtherServers As New List(Of String)()
        End Class

        ''' <summary>
        ''' One probe round. A typed address is probed alone because the operator said
        ''' where to look; an empty one runs the network scan. The first candidate that
        ''' hosts a complete BTA_OSG_DB wins, and every failure class an operator meets
        ''' gets its own outcome so the dialog can answer with the right next action.
        ''' </summary>
        Private Shared Function ProbeTargets(typedServer As String, typedPort As String, user As String, pass As String) As ConnectResult
            Dim results As New List(Of SqlInstanceDiscovery.ScanResult)()
            If typedServer.Length > 0 Then
                Dim portValue As Integer? = Nothing
                If Not typedServer.Contains("\"c) Then
                    Dim parsed As Integer
                    If Integer.TryParse(typedPort.Trim(), parsed) Then portValue = parsed
                End If
                results.Add(New SqlInstanceDiscovery.ScanResult With {
                            .Server = typedServer,
                            .Port = portValue,
                            .Result = SqlInstanceDiscovery.Probe(typedServer, portValue, user, pass)})
            Else
                results = SqlInstanceDiscovery.Discover(user, pass)
            End If

            Dim found = results.FindAll(Function(x) x.Result = SqlInstanceDiscovery.ProbeResult.FoundDatabase)
            Dim rejected = results.FindAll(Function(x) x.Result = SqlInstanceDiscovery.ProbeResult.LoginRejected)
            Dim missingDb = results.FindAll(Function(x) x.Result = SqlInstanceDiscovery.ProbeResult.NoDatabase)

            If found.Count > 0 Then
                Dim best = found(0)
                Dim others As New List(Of String)()
                For i As Integer = 1 To found.Count - 1
                    others.Add(found(i).Server)
                Next
                Dim portText = If(best.Port.HasValue, best.Port.Value.ToString(), "")
                Dim complete As Boolean = SchemaLadderComplete(BuildConnectionString(best.Server, portText, user, pass))
                Return New ConnectResult With {
                    .Outcome = If(complete, ConnectOutcome.Connected, ConnectOutcome.SchemaIncomplete),
                    .Server = best.Server,
                    .Port = best.Port,
                    .OtherServers = others
                }
            End If
            If rejected.Count > 0 Then
                Return New ConnectResult With {.Outcome = ConnectOutcome.Rejected, .Server = rejected(0).Server, .Port = rejected(0).Port}
            End If
            If missingDb.Count > 0 Then
                Return New ConnectResult With {.Outcome = ConnectOutcome.MissingDatabase, .Server = missingDb(0).Server, .Port = missingDb(0).Port}
            End If
            Return New ConnectResult With {.Outcome = ConnectOutcome.Unreachable}
        End Function

        ''' <summary>
        ''' The migration tail of 008-010: a database that stops mid-ladder must read as
        ''' incomplete so Set up this server is offered instead of a broken seat.
        ''' </summary>
        Private Shared Function SchemaLadderComplete(connectionString As String) As Boolean
            Try
                Using conn As New SqlConnection(connectionString)
                    conn.Open()
                    Using cmd As New SqlCommand(
                        "SELECT CASE WHEN COL_LENGTH('dbo.tbl_Documents', 'FlowDirection') IS NOT NULL " &
                        "AND COL_LENGTH('dbo.tbl_Documents', 'ExternalControlNumber') IS NOT NULL " &
                        "AND COL_LENGTH('dbo.tbl_Users', 'CanRoute') IS NOT NULL " &
                        "AND COL_LENGTH('dbo.tbl_AuditTrail', 'RowHash') IS NOT NULL " &
                        "AND COLUMNPROPERTY(OBJECT_ID('dbo.tbl_RoutingLogs'), 'ToStatusID', 'AllowsNull') = 1 THEN 1 ELSE 0 END", conn)
                        Return Convert.ToInt32(cmd.ExecuteScalar()) = 1
                    End Using
                End Using
            Catch ex As Exception
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Blank password on a configured station reuses the saved sign-in: the
        ''' SYSADMIN gate already ran to open this dialog, and re-typing an unchanged
        ''' password is friction. A malformed saved string falls through to the
        ''' typed-password rule.
        ''' </summary>
        Private Function EffectiveCredentials() As (User As String, Password As String)
            If txtPassword.Text.Length > 0 Then Return (AppLoginName, txtPassword.Text)
            If AppSettings.Instance.IsConfigured AndAlso AppSettings.Instance.DatabaseSettings.UseSqlServer Then
                Try
                    Dim saved As New SqlConnectionStringBuilder(AppSettings.Instance.DatabaseSettings.ConnectionString)
                    If saved.Password.Length > 0 Then
                        Return (If(saved.UserID.Length > 0, saved.UserID, AppLoginName), saved.Password)
                    End If
                Catch ex As Exception
                End Try
            End If
            Return (AppLoginName, "")
        End Function

        Private Sub OfferPrepare(result As ConnectResult, message As String)
            _probeServer = result.Server
            Dim localNames = SqlInstanceDiscovery.LocalInstanceNames()
            Dim isLocal As Boolean = _probeServer.Equals(Environment.MachineName, StringComparison.OrdinalIgnoreCase) OrElse
                                     _probeServer.StartsWith(Environment.MachineName & "\", StringComparison.OrdinalIgnoreCase)
            If isLocal AndAlso localNames.Count > 0 Then
                btnPrepare.Visible = True
                SetStatus(message & " You can set up this server below.", CivicCalmTheme.ColorWarning)
            Else
                btnPrepare.Visible = False
                SetStatus(message, CivicCalmTheme.ColorWarning)
            End If
        End Sub

        Private Sub SaveSqlServerSettings(server As String, port As Integer?, user As String, pass As String)
            Dim s = AppSettings.Instance
            s.DatabaseSettings.UseSqlServer = True
            s.DatabaseSettings.ConnectionString = BuildConnectionString(server, If(port.HasValue, port.Value.ToString(), Nothing), user, pass)
            s.PortalSettings.PortalEnabled = chkPortal.Checked
            s.AnalyticsEnabled = chkAnalytics.Checked
            If chkPortal.Checked Then
                s.PortalSettings.BaseUrl = LocalPortalBaseUrl
                ' An empty (or the published demo default) key would leave the intake API
                ' open to whoever reads the repo, so a fresh 32-byte value generates here
                ' and travels to the portal process through its environment automatically.
                If String.IsNullOrWhiteSpace(s.PortalSettings.BridgeKey) OrElse String.Equals(s.PortalSettings.BridgeKey, PublishedDemoBridgeKey, StringComparison.OrdinalIgnoreCase) Then
                    s.PortalSettings.BridgeKey = Convert.ToHexString(System.Security.Cryptography.RandomNumberGenerator.GetBytes(32)).ToLowerInvariant()
                End If
            End If
            AppSettings.Save(s)
            ShowSavedMessage(server)
            Me.DialogResult = DialogResult.OK
            Me.Close()
        End Sub

        ''' <summary>
        ''' One confirmation per saved change, because the caller cannot tell which of its own
        ''' settings moved. The mode applies to the running session as soon as the caller
        ''' reloads; the portal and Data Analytics switches are read while the main form builds
        ''' its navigation and pollers, so those name the next launch instead of lying about it.
        ''' </summary>
        Private Sub ShowSavedMessage(server As String)
            Dim s = AppSettings.Instance
            Dim restart = ""
            If s.PortalSettings.PortalEnabled <> _startedPortal OrElse s.AnalyticsEnabled <> _startedAnalytics Then
                restart = vbCrLf & vbCrLf & "Restart the application to apply the portal and Data Analytics changes."
            End If

            Dim headline As String
            If _justPrepared Then
                headline = "Server ready. The next screen names the System Administrator and enrols that badge; no other card can sign in."
            ElseIf Not s.DatabaseSettings.UseSqlServer Then
                headline = "This station is now in demo mode. It works from its local cache and does not touch the office server."
            ElseIf _switchingFromDemo Then
                headline = "Connected to " & server & ". This station is now in office mode and reads and writes the office database."
            Else
                headline = "Workstation configuration saved."
            End If
            _justPrepared = False
            _switchingFromDemo = False
            MessageBox.Show(headline & restart, "Station Setup", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub

        Private Sub OnDemoLink(sender As Object, e As LinkLabelLinkClickedEventArgs)
            If AppSettings.Instance.IsConfigured Then
                If MessageBox.Show("Switch this station to demo mode? It will stop using the office server until you connect again.",
                                   "Run Demo Mode", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) <> DialogResult.Yes Then Return
            End If
            Dim s = AppSettings.Instance
            s.DatabaseSettings.UseSqlServer = False
            s.PortalSettings.PortalEnabled = True
            s.PortalSettings.BaseUrl = LocalPortalBaseUrl
            If String.IsNullOrWhiteSpace(s.PortalSettings.BridgeKey) OrElse String.Equals(s.PortalSettings.BridgeKey, PublishedDemoBridgeKey, StringComparison.OrdinalIgnoreCase) Then
                s.PortalSettings.BridgeKey = Convert.ToHexString(System.Security.Cryptography.RandomNumberGenerator.GetBytes(32)).ToLowerInvariant()
            End If
            AppSettings.Save(s)
            ShowSavedMessage("")
            Me.DialogResult = DialogResult.OK
            Me.Close()
        End Sub

        Private Async Sub OnPrepareServer(sender As Object, e As EventArgs)
            Dim credentials = EffectiveCredentials()
            If credentials.Password.Length = 0 Then
                SetStatus("Type the shared office password first; server setup creates the login with it.", CivicCalmTheme.ColorWarning)
                txtPassword.Focus()
                Return
            End If
            Dim confirmText =
                "Set up this computer as the office server?" & vbCrLf & vbCrLf &
                "With the current Windows account it will:" & vbCrLf &
                "1. Create the OSG document database and its tables. Existing data is never dropped." & vbCrLf &
                "2. Create the shared office sign-in the seats use (bta_app), or reset its password to the one you typed." & vbCrLf &
                "3. Allow other computers on the network to reach SQL Server." & vbCrLf &
                "4. Restart the SQL Server service and open the two network ports in Windows Firewall." & vbCrLf &
                "5. Test the connection again." & vbCrLf & vbCrLf &
                "Step 4 needs administrator permission; Windows will ask once. Continue?"
            If MessageBox.Show(confirmText, "Set up this server", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) <> DialogResult.Yes Then Return

            SetBusy(True)
            SetStatus("Setting up the server: database, schema, login, and network settings...", CivicCalmTheme.ColorInfo)
            ' The prepare target is always the local candidate discovery answered with,
            ' and the admin connection rides shared memory, so it works even while TCP
            ' is still disabled on a fresh Express install.
            Dim adminConnStr = BuildAdminConnectionString(_probeServer, "")
            Dim prepError As String = ""
            Await Task.Run(Sub()
                               Try
                                   DatabaseProvisioner.Provision(adminConnStr)
                                   DatabaseProvisioner.EnsureBtaAppLogin(adminConnStr, AppLoginName, credentials.Password)
                                   DatabaseProvisioner.EnsureTcpEnabled(adminConnStr)
                               Catch ex As Exception
                                   prepError = ex.Message
                               End Try
                           End Sub)
            If prepError.Length > 0 Then
                SetBusy(False)
                SetStatus("Setup failed: " & prepError & " Run this on the server PC as a Windows administrator.", CivicCalmTheme.ColorDanger)
                Return
            End If

            SetStatus("Restarting SQL Server and opening the firewall; Windows will ask for permission once...", CivicCalmTheme.ColorInfo)
            Dim netCode As Integer = 1
            Dim netError As String = ""
            Try
                netCode = Await Task.Run(Function() NetworkPrep.RunElevated())
            Catch ex As Exception
                netError = ex.Message
            End Try
            If netCode <> 0 Then
                SetBusy(False)
                SetStatus("The database is ready, but the network steps did not finish" & If(netError.Length > 0, " (" & netError & ")", "") & ". See the deployment guide for the firewall and service steps, then connect again.", CivicCalmTheme.ColorWarning)
                Return
            End If

            ' The SQL service has restarted; give the TCP listener a moment before the
            ' re-probe, whose 2s connect timeout is shorter than a cold service start.
            Await Task.Delay(3000)
            SetStatus("Setup finished. Re-testing the connection...", CivicCalmTheme.ColorInfo)
            _justPrepared = True
            Await RunConnectAsync()
        End Sub

        Private Sub SetStatus(text As String, color As Color)
            lblStatus.Text = text
            lblStatus.ForeColor = color
        End Sub

        Private Sub SetBusy(busy As Boolean)
            _busy = busy
            btnConnect.Enabled = Not busy
            btnPrepare.Enabled = Not busy
            txtPassword.Enabled = Not busy
            lnkDemo.Enabled = Not busy
            lnkOptions.Enabled = Not busy
        End Sub
    End Class
End Namespace
