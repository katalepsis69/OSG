Option Explicit On
Option Strict On

Imports System
Imports System.Drawing
Imports System.Windows.Forms

Namespace BTA_OSG
    Public Class FormMoveStorage
        Inherits Form

        Private DocID As Integer
        Private StaffName As String
        Private txtFromLoc As TextBox
        Private txtToLoc As TextBox
        Private txtReason As TextBox
        Private btnSave As Button
        Private btnCancel As Button

        Public Sub New(docId As Integer, currentLoc As String, user As String)
            Me.DocID = docId
            Me.StaffName = user

            Me.Text = "Transfer Physical Landmark Storage Location"
            Me.Size = New Size(500, 320)
            Me.MinimumSize = New Size(470, 300)
            Me.StartPosition = FormStartPosition.CenterParent
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.BackColor = CivicCalmTheme.ColorCanvas
            Me.Font = CivicCalmTheme.FontBody

            Dim pnlCard As New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(20),
                .Margin = New Padding(16)
            }

            Dim tblLayout As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 2,
                .RowCount = 4
            }
            tblLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 150.0F))
            tblLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            tblLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 40.0F))
            tblLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 40.0F))
            tblLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 60.0F))
            tblLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 50.0F))

            txtFromLoc = New TextBox With {.Dock = DockStyle.Fill, .Text = currentLoc, .ReadOnly = True, .BackColor = CivicCalmTheme.ColorWell, .ForeColor = CivicCalmTheme.ColorInkMuted, .BorderStyle = BorderStyle.FixedSingle}
            txtToLoc = New TextBox With {.Dock = DockStyle.Fill, .Text = "CAB-A/S-3/BOX-05", .BackColor = CivicCalmTheme.ColorSurface, .ForeColor = CivicCalmTheme.ColorInk, .BorderStyle = BorderStyle.FixedSingle}
            txtReason = New TextBox With {.Dock = DockStyle.Fill, .Multiline = True, .Text = "Archival reorganization per OSG directive.", .BackColor = CivicCalmTheme.ColorSurface, .ForeColor = CivicCalmTheme.ColorInk, .BorderStyle = BorderStyle.FixedSingle}

            Dim flwButtons As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.RightToLeft,
                .WrapContents = False
            }

            btnCancel = New Button With {
                .Text = "&Cancel",
                .Size = New Size(95, 34),
                .BackColor = CivicCalmTheme.ColorWell,
                .ForeColor = CivicCalmTheme.ColorInk,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(8, 6, 0, 0)
            }
            btnCancel.FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder
            AddHandler btnCancel.Click, Sub()
                                            Me.DialogResult = DialogResult.Cancel
                                            Me.Close()
                                        End Sub

            btnSave = New Button With {
                .Text = "&Transfer Storage",
                .Size = New Size(160, 34),
                .BackColor = CivicCalmTheme.ColorPrimary,
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(0, 6, 0, 0)
            }
            btnSave.FlatAppearance.BorderSize = 0
            AddHandler btnSave.Click, AddressOf OnSave

            flwButtons.Controls.AddRange(New Control() {btnCancel, btnSave})

            tblLayout.Controls.Add(New Label With {.Text = "Current Location:", .Font = CivicCalmTheme.FontFieldLabel, .ForeColor = CivicCalmTheme.ColorInkMuted, .Dock = DockStyle.Fill, .TextAlign = ContentAlignment.MiddleLeft}, 0, 0)
            tblLayout.Controls.Add(txtFromLoc, 1, 0)
            tblLayout.Controls.Add(New Label With {.Text = "New (Cab/Shelf/Box):", .Font = CivicCalmTheme.FontFieldLabel, .ForeColor = CivicCalmTheme.ColorInkMuted, .Dock = DockStyle.Fill, .TextAlign = ContentAlignment.MiddleLeft}, 0, 1)
            tblLayout.Controls.Add(txtToLoc, 1, 1)
            tblLayout.Controls.Add(New Label With {.Text = "Reason for Move:", .Font = CivicCalmTheme.FontFieldLabel, .ForeColor = CivicCalmTheme.ColorInkMuted, .Dock = DockStyle.Fill, .TextAlign = ContentAlignment.TopLeft}, 0, 2)
            tblLayout.Controls.Add(txtReason, 1, 2)
            tblLayout.Controls.Add(flwButtons, 1, 3)

            pnlCard.Controls.Add(tblLayout)
            Me.Controls.Add(pnlCard)

            Me.AcceptButton = btnSave
            Me.CancelButton = btnCancel
        End Sub

        Private Sub OnSave(sender As Object, e As EventArgs)
            EmbeddedDB.AddMovementLog(DocID, txtFromLoc.Text.Trim(), txtToLoc.Text.Trim(), StaffName, txtReason.Text.Trim())
            EmbeddedDB.LogAudit(StaffName, String.Format("Transferred Doc #{0} physical location from {1} to {2}", DocID, txtFromLoc.Text, txtToLoc.Text))
            Me.DialogResult = DialogResult.OK
            Me.Close()
        End Sub
    End Class
End Namespace
