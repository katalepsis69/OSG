Imports System.Drawing
Imports System.Windows.Forms


Namespace BTA_OSG
    Public Class FormLogin
        Inherits Form

        Private txtScannerInput As TextBox
        Private lblStatus As Label
        Private lblTitle As Label
        Private btnCancel As Button
        Private pnlSim As Panel

        <System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)>
        Public Property AuthenticatedSession As SessionContext

        Public Sub New()
            InitializeComponent()
            ApplyTheme()
        End Sub

        Private Sub InitializeComponent()
            Me.Size = New Size(400, 300)
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.Text = "BTA OSG - Login"

            lblTitle = New Label()
            lblTitle.Text = "RFID TAP BADGE SCANNER"
            lblTitle.Font = New Font("Segoe UI", 14, FontStyle.Bold)
            lblTitle.TextAlign = ContentAlignment.MiddleCenter
            lblTitle.Dock = DockStyle.Top
            lblTitle.Height = 60

            lblStatus = New Label()
            lblStatus.Text = "Ready to Scan..."
            lblStatus.Font = New Font("Segoe UI", 12)
            lblStatus.TextAlign = ContentAlignment.MiddleCenter
            lblStatus.Dock = DockStyle.Fill

            txtScannerInput = New TextBox()
            txtScannerInput.Location = New Point(-100, -100) ' Hidden
            AddHandler txtScannerInput.KeyDown, AddressOf OnScannerInput

            btnCancel = New Button()
            btnCancel.Text = "Cancel"
            btnCancel.Dock = DockStyle.Bottom
            btnCancel.Height = 40
            AddHandler btnCancel.Click, Sub(s, e) Me.DialogResult = DialogResult.Cancel

            pnlSim = New Panel()
            pnlSim.Dock = DockStyle.Bottom
            pnlSim.Height = 50

            Dim btnSimSg = New Button() With {.Text = "Sim SG", .Width = 100, .Dock = DockStyle.Left}
            Dim btnSimAdmin = New Button() With {.Text = "Sim Admin", .Width = 100, .Dock = DockStyle.Left}
            Dim btnSimStaff = New Button() With {.Text = "Sim Staff", .Width = 100, .Dock = DockStyle.Left}
            
            AddHandler btnSimSg.Click, Sub(s, e) ProcessRfid("SG_CARD_001")
            AddHandler btnSimAdmin.Click, Sub(s, e) ProcessRfid("ADMIN_CARD_001")
            AddHandler btnSimStaff.Click, Sub(s, e) ProcessRfid("STAFF_CARD_001")

            pnlSim.Controls.Add(btnSimStaff)
            pnlSim.Controls.Add(btnSimAdmin)
            pnlSim.Controls.Add(btnSimSg)

            Me.Controls.Add(txtScannerInput)
            Me.Controls.Add(lblStatus)
            If AppStartup.Settings.RfidSettings.SimulatorEnabled Then
                Me.Controls.Add(pnlSim)
            End If
            Me.Controls.Add(btnCancel)
            Me.Controls.Add(lblTitle)
        End Sub

        Private Sub ApplyTheme()
            Me.BackColor = Color.FromArgb(15, 23, 42)
            lblTitle.ForeColor = Color.FromArgb(79, 70, 229)
            lblStatus.ForeColor = Color.FromArgb(203, 213, 225)
            
            btnCancel.FlatStyle = FlatStyle.Flat
            btnCancel.BackColor = Color.FromArgb(30, 41, 59)
            btnCancel.ForeColor = Color.FromArgb(248, 250, 252)
        End Sub

        Private Sub OnScannerInput(sender As Object, e As KeyEventArgs)
            If e.KeyCode = Keys.Enter Then
                Dim rfid = txtScannerInput.Text.Trim()
                txtScannerInput.Clear()
                ProcessRfid(rfid)
            End If
        End Sub

        Private Sub ProcessRfid(rfid As String)
            Try
                Dim session = AppStartup.AuthService.AuthenticateByCard(rfid)
                If session IsNot Nothing Then
                    AuthenticatedSession = session
                    Me.DialogResult = DialogResult.OK
                Else
                    lblStatus.Text = "Invalid Card. Try again."
                    lblStatus.ForeColor = Color.FromArgb(239, 68, 68)
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Login Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Protected Overrides Sub OnShown(e As EventArgs)
            MyBase.OnShown(e)
            txtScannerInput.Focus()
        End Sub
    End Class
End Namespace
