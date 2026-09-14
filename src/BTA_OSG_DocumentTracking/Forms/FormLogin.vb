Imports System
Imports System.Drawing
Imports System.Windows.Forms

Namespace BTA_OSG
    Public Class FormLogin
        Inherits Form

        <System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)>
        Public Property ScannedUID As String = ""
        <System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)>
        Public Property AuthenticatedSession As SessionContext

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
            Dim btn As New Button With {
                .Text = text,
                .Location = loc,
                .Size = New Size(190, 36),
                .BackColor = bg,
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
                .Cursor = Cursors.Hand
            }
            btn.FlatAppearance.BorderSize = 0
            Return btn
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
            ScannedUID = uid.Trim().ToUpperInvariant()
            Me.DialogResult = DialogResult.OK
            Me.Close()
        End Sub
    End Class
End Namespace
