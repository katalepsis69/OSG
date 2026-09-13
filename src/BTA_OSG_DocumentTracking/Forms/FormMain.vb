Imports System.Drawing
Imports System.Windows.Forms


Namespace BTA_OSG
    Public Class FormMain
        Inherits Form

        Private pnlSidebar As Panel
        Private pnlHeader As Panel
        Private pnlContent As Panel
        Private lblUserBadge As Label
        Private tmrSession As Timer

        ' View Panels
        Private viewDashboard As Panel
        Private viewDocReg As Panel
        Private viewDirectives As Panel
        Private viewSearch As Panel
        Private viewUserAdmin As Panel
        Private viewAudit As Panel

        Public Sub New()
            InitializeComponent()
            ApplyTheme()
            SetupSessionTimer()
            ShowLogin()
        End Sub

        Private Sub InitializeComponent()
            Me.Size = New Size(1380, 850)
            Me.MinimumSize = New Size(1100, 720)
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.Text = "BTA OSG Document Tracking"

            ' Sidebar
            pnlSidebar = New Panel() With {.Width = 240, .Dock = DockStyle.Left}
            CreateNavButton("Dashboard", AddressOf ShowDashboard)
            CreateNavButton("Document Registry", AddressOf ShowDocReg)
            CreateNavButton("SG Directives", AddressOf ShowDirectives)
            CreateNavButton("Search && Storage", AddressOf ShowSearch)
            CreateNavButton("User && RFID Admin", AddressOf ShowUserAdmin)
            CreateNavButton("Audit Trail", AddressOf ShowAudit)

            ' Header
            pnlHeader = New Panel() With {.Height = 70, .Dock = DockStyle.Top}
            Dim lblTitle = New Label() With {.Text = "BTA OSG DOC TRACKER", .Font = New Font("Segoe UI", 16, FontStyle.Bold), .Dock = DockStyle.Left, .Width = 300, .TextAlign = ContentAlignment.MiddleLeft}
            
            lblUserBadge = New Label() With {.Text = "Not Logged In", .Dock = DockStyle.Right, .Width = 200, .TextAlign = ContentAlignment.MiddleRight}
            Dim btnLogout = New Button() With {.Text = "Logout", .Dock = DockStyle.Right, .Width = 100, .FlatStyle = FlatStyle.Flat}
            AddHandler btnLogout.Click, Sub() ShowLogin()

            pnlHeader.Controls.Add(lblTitle)
            pnlHeader.Controls.Add(lblUserBadge)
            pnlHeader.Controls.Add(btnLogout)

            ' Content Area
            pnlContent = New Panel() With {.Dock = DockStyle.Fill}
            
            Me.Controls.Add(pnlContent)
            Me.Controls.Add(pnlSidebar)
            Me.Controls.Add(pnlHeader)

            InitializeViews()
        End Sub

        Private Sub CreateNavButton(text As String, handler As EventHandler)
            Dim btn = New Button() With {.Text = text, .Dock = DockStyle.Top, .Height = 50, .FlatStyle = FlatStyle.Flat, .TextAlign = ContentAlignment.MiddleLeft, .Padding = New Padding(20, 0, 0, 0)}
            AddHandler btn.Click, handler
            pnlSidebar.Controls.Add(btn)
            pnlSidebar.Controls.SetChildIndex(btn, 0)
        End Sub

        Private Sub InitializeViews()
            viewDashboard = New Panel() With {.Dock = DockStyle.Fill, .Visible = False}
            viewDashboard.Controls.Add(New Label() With {.Text = "Dashboard View", .ForeColor = Color.White})

            viewDocReg = New Panel() With {.Dock = DockStyle.Fill, .Visible = False}
            viewDocReg.Controls.Add(New Label() With {.Text = "Document Registration View", .ForeColor = Color.White})

            viewDirectives = New Panel() With {.Dock = DockStyle.Fill, .Visible = False}
            viewDirectives.Controls.Add(New Label() With {.Text = "SG Directives View", .ForeColor = Color.White})

            viewSearch = New Panel() With {.Dock = DockStyle.Fill, .Visible = False}
            viewSearch.Controls.Add(New Label() With {.Text = "Search & Storage View", .ForeColor = Color.White})

            viewUserAdmin = New Panel() With {.Dock = DockStyle.Fill, .Visible = False}
            viewUserAdmin.Controls.Add(New Label() With {.Text = "User Admin View", .ForeColor = Color.White})

            viewAudit = New Panel() With {.Dock = DockStyle.Fill, .Visible = False}
            viewAudit.Controls.Add(New Label() With {.Text = "Audit Trail View", .ForeColor = Color.White})

            pnlContent.Controls.Add(viewDashboard)
            pnlContent.Controls.Add(viewDocReg)
            pnlContent.Controls.Add(viewDirectives)
            pnlContent.Controls.Add(viewSearch)
            pnlContent.Controls.Add(viewUserAdmin)
            pnlContent.Controls.Add(viewAudit)
        End Sub

        Private Sub ApplyTheme()
            Me.BackColor = Color.FromArgb(15, 23, 42)
            pnlSidebar.BackColor = Color.FromArgb(30, 41, 59)
            pnlHeader.BackColor = Color.FromArgb(30, 41, 59)
            lblUserBadge.ForeColor = Color.FromArgb(203, 213, 225)
            
            For Each c As Control In pnlSidebar.Controls
                If TypeOf c Is Button Then
                    c.BackColor = Color.FromArgb(30, 41, 59)
                    c.ForeColor = Color.FromArgb(203, 213, 225)
                End If
            Next
        End Sub

        Private Sub SetupSessionTimer()
            tmrSession = New Timer() With {.Interval = 30000}
            AddHandler tmrSession.Tick, Sub(s, e)
                If SessionManager.IsLoggedIn AndAlso SessionManager.CurrentSession.IsExpired(AppStartup.Settings.SessionSettings.TimeoutMinutes) Then
                    tmrSession.Stop()
                    MessageBox.Show("Session expired due to inactivity.", "Session Timeout", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    ShowLogin()
                End If
            End Sub
        End Sub

        Private Sub ShowLogin()
            SessionManager.EndSession()
            Me.Hide()
            Dim login = New FormLogin()
            If login.ShowDialog() = DialogResult.OK Then
                SessionManager.SetCurrentSession(login.AuthenticatedSession)
                UpdateHeaderBadge()
                tmrSession.Start()
                Me.Show()
                ShowDashboard(Nothing, Nothing)
            Else
                Application.Exit()
            End If
        End Sub

        Private Sub UpdateHeaderBadge()
            If SessionManager.IsLoggedIn Then
                Dim roleStr As String = ""
                If SessionManager.CurrentSession.Roles IsNot Nothing AndAlso SessionManager.CurrentSession.Roles.Count > 0 Then
                    roleStr = " (" & SessionManager.CurrentSession.Roles(0).RoleName & ")"
                End If
                lblUserBadge.Text = SessionManager.CurrentSession.User.Username & roleStr
            End If
        End Sub

        Private Sub HideAllViews()
            viewDashboard.Visible = False
            viewDocReg.Visible = False
            viewDirectives.Visible = False
            viewSearch.Visible = False
            viewUserAdmin.Visible = False
            viewAudit.Visible = False
        End Sub

        Private Sub ShowDashboard(sender As Object, e As EventArgs)
            HideAllViews()
            viewDashboard.Visible = True
        End Sub

        Private Sub ShowDocReg(sender As Object, e As EventArgs)
            HideAllViews()
            viewDocReg.Visible = True
        End Sub

        Private Sub ShowDirectives(sender As Object, e As EventArgs)
            If Not SessionManager.HasPermission(RbacPolicy.DIRECTIVE_ISSUE) Then
                MessageBox.Show("Access Denied.", "Security", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If
            HideAllViews()
            viewDirectives.Visible = True
        End Sub

        Private Sub ShowSearch(sender As Object, e As EventArgs)
            HideAllViews()
            viewSearch.Visible = True
        End Sub

        Private Sub ShowUserAdmin(sender As Object, e As EventArgs)
            If Not SessionManager.HasPermission(RbacPolicy.USER_MANAGE) Then
                MessageBox.Show("Access Denied.", "Security", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If
            HideAllViews()
            viewUserAdmin.Visible = True
        End Sub

        Private Sub ShowAudit(sender As Object, e As EventArgs)
            If Not SessionManager.HasPermission(RbacPolicy.AUDIT_VIEW) Then
                MessageBox.Show("Access Denied.", "Security", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If
            HideAllViews()
            viewAudit.Visible = True
        End Sub
    End Class
End Namespace
