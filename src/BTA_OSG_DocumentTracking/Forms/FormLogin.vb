Option Explicit On
Option Strict On

Imports System
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms

Namespace BTA_OSG
    Public Class FormLogin
        Inherits Form

        <System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)>
        Public Property ScannedUID As String = ""

        Private lblScanStatus As Label
        Private ReadOnly _rfidBuffer As New StringBuilder()
        Private btnCancel As Button

        Public Sub New()
            InitializeForm()
        End Sub

        Private Sub InitializeForm()
            Me.Text = "OSG RFID Smart Card Authentication Tap Scanner"
            Me.Size = New Size(500, 380)
            Me.MinimumSize = New Size(480, 360)
            Me.StartPosition = FormStartPosition.CenterParent
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.BackColor = CivicCalmTheme.ColorCanvas
            Me.Font = CivicCalmTheme.FontBody
            Me.KeyPreview = True

            Dim pnlOuter As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .RowCount = 3,
                .Padding = New Padding(20)
            }
            pnlOuter.RowStyles.Add(New RowStyle(SizeType.Absolute, 140.0F))
            pnlOuter.RowStyles.Add(New RowStyle(SizeType.Absolute, 32.0F))
            pnlOuter.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

            ' Top Card: Scanner instructions and status
            Dim pnlCard As New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(16),
                .Margin = New Padding(0, 0, 0, 12)
            }

            Dim lblIcon As New Label With {
                .Text = "RFID BADGE SCANNER : ACTIVE",
                .ForeColor = CivicCalmTheme.ColorPrimary,
                .Font = CivicCalmTheme.FontSectionHeader,
                .Dock = DockStyle.Top,
                .Height = 24
            }

            lblScanStatus = New Label With {
                .Text = "Ready to scan. Please tap your physical RFID card on the USB reader." & vbCrLf & "The terminal will capture your badge ID automatically.",
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .Font = CivicCalmTheme.FontBody,
                .Dock = DockStyle.Fill,
                .TextAlign = ContentAlignment.MiddleLeft
            }

            pnlCard.Controls.Add(lblScanStatus)
            pnlCard.Controls.Add(lblIcon)

            ' Middle Label
            Dim lblSim As New Label With {
                .Text = "Or select a test badge to simulate tap:",
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .Font = CivicCalmTheme.FontMicrocopy,
                .Dock = DockStyle.Fill,
                .TextAlign = ContentAlignment.BottomLeft
            }

            ' Bottom Grid: Simulation & Cancel buttons
            Dim tblButtons As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 2,
                .RowCount = 2,
                .Margin = New Padding(0, 6, 0, 0)
            }
            tblButtons.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0F))
            tblButtons.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0F))
            tblButtons.RowStyles.Add(New RowStyle(SizeType.Percent, 50.0F))
            tblButtons.RowStyles.Add(New RowStyle(SizeType.Percent, 50.0F))

            Dim btnSG = CreateSimButton("&SG Card (88A9F321)", CivicCalmTheme.ColorPrimary)
            Dim btnAdmin = CreateSimButton("&Admin Card (77C3D987)", ColorTranslator.FromHtml("#0284C7"))
            Dim btnStaff = CreateSimButton("S&taff Card (55E5F666)", ColorTranslator.FromHtml("#D97706"))
            btnCancel = New Button With {
                .Text = "&Cancel",
                .Dock = DockStyle.Fill,
                .Margin = New Padding(4),
                .BackColor = CivicCalmTheme.ColorWell,
                .ForeColor = CivicCalmTheme.ColorInk,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand
            }
            btnCancel.FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder

            AddHandler btnSG.Click, Sub() SubmitUID("88A9F321")
            AddHandler btnAdmin.Click, Sub() SubmitUID("77C3D987")
            AddHandler btnStaff.Click, Sub() SubmitUID("55E5F666")
            AddHandler btnCancel.Click, Sub()
                                            Me.DialogResult = DialogResult.Cancel
                                            Me.Close()
                                        End Sub

            Me.CancelButton = btnCancel

            tblButtons.Controls.Add(btnSG, 0, 0)
            tblButtons.Controls.Add(btnAdmin, 1, 0)
            tblButtons.Controls.Add(btnStaff, 0, 1)
            tblButtons.Controls.Add(btnCancel, 1, 1)

            pnlOuter.Controls.Add(pnlCard, 0, 0)
            pnlOuter.Controls.Add(lblSim, 0, 1)
            pnlOuter.Controls.Add(tblButtons, 0, 2)

            Me.Controls.Add(pnlOuter)

            AddHandler Me.KeyPress, AddressOf OnFormKeyPress
        End Sub

        Private Function CreateSimButton(text As String, bg As Color) As Button
            Dim btn As New Button With {
                .Text = text,
                .Dock = DockStyle.Fill,
                .Margin = New Padding(4),
                .BackColor = bg,
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand
            }
            btn.FlatAppearance.BorderSize = 0
            Return btn
        End Function

        Private Sub OnFormKeyPress(sender As Object, e As KeyPressEventArgs)
            If e.KeyChar = ChrW(Keys.Enter) OrElse e.KeyChar = vbCr OrElse e.KeyChar = vbLf Then
                e.Handled = True
                Dim uid = _rfidBuffer.ToString().Trim()
                _rfidBuffer.Clear()
                If Not String.IsNullOrEmpty(uid) Then
                    SubmitUID(uid)
                End If
            ElseIf Not Char.IsControl(e.KeyChar) Then
                _rfidBuffer.Append(e.KeyChar)
            End If
        End Sub

        Private Sub SubmitUID(uid As String)
            ScannedUID = uid.Trim().ToUpperInvariant()
            Me.DialogResult = DialogResult.OK
            Me.Close()
        End Sub
    End Class
End Namespace
