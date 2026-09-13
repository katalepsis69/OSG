Imports System.Drawing
Imports System.Windows.Forms


Namespace BTA_OSG
    Public Class FormSettings
        Inherits Form

        Public Sub New()
            InitializeComponent()
            ApplyTheme()
            LoadSettings()
        End Sub

        Private txtConnString As TextBox
        Private chkSim As CheckBox
        Private lblTimeout As Label
        Private btnClose As Button

        Private Sub InitializeComponent()
            Me.Size = New Size(400, 250)
            Me.StartPosition = FormStartPosition.CenterParent
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.Text = "System Settings"

            Dim layout = New TableLayoutPanel()
            layout.Dock = DockStyle.Fill
            layout.ColumnCount = 2
            layout.RowCount = 4

            layout.Controls.Add(New Label() With {.Text = "DB Connection:", .ForeColor = Color.White, .AutoSize = True}, 0, 0)
            txtConnString = New TextBox() With {.Dock = DockStyle.Fill, .ReadOnly = True}
            layout.Controls.Add(txtConnString, 1, 0)

            layout.Controls.Add(New Label() With {.Text = "RFID Simulator:", .ForeColor = Color.White, .AutoSize = True}, 0, 1)
            chkSim = New CheckBox() With {.Enabled = False}
            layout.Controls.Add(chkSim, 1, 1)

            layout.Controls.Add(New Label() With {.Text = "Session Timeout:", .ForeColor = Color.White, .AutoSize = True}, 0, 2)
            lblTimeout = New Label() With {.ForeColor = Color.White, .AutoSize = True}
            layout.Controls.Add(lblTimeout, 1, 2)

            btnClose = New Button() With {.Text = "Close", .Dock = DockStyle.Bottom, .Height = 40, .FlatStyle = FlatStyle.Flat}
            AddHandler btnClose.Click, Sub(s, e) Me.Close()

            Me.Controls.Add(layout)
            Me.Controls.Add(btnClose)
        End Sub

        Private Sub ApplyTheme()
            Me.BackColor = Color.FromArgb(15, 23, 42)
            btnClose.BackColor = Color.FromArgb(30, 41, 59)
            btnClose.ForeColor = Color.White
        End Sub

        Private Sub LoadSettings()
            If AppStartup.Settings IsNot Nothing Then
                txtConnString.Text = AppStartup.Settings.DatabaseSettings.ConnectionString
                chkSim.Checked = AppStartup.Settings.RfidSettings.SimulatorEnabled
                lblTimeout.Text = AppStartup.Settings.SessionSettings.TimeoutMinutes.ToString() & " minutes"
            End If
        End Sub
    End Class
End Namespace
