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
        Private StaffUserId As Integer
        Private txtFromLoc As TextBox
        Private txtToLoc As TextBox
        Private txtReason As TextBox
        Private btnSave As Button
        Private btnCancel As Button

        Public Sub New(docId As Integer, currentLoc As String, user As String, Optional staffUserId As Integer = 1)
            Me.DocID = docId
            Me.StaffName = user
            Me.StaffUserId = staffUserId
            AppAssets.ApplyFormIcon(Me)

            Me.Text = "Transfer Landmark Storage Location"
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

            txtFromLoc = New TextBox With {.Dock = DockStyle.Fill, .Text = currentLoc, .ReadOnly = True, .TabIndex = 1, .BackColor = CivicCalmTheme.ColorWell, .ForeColor = CivicCalmTheme.ColorInkMuted, .BorderStyle = BorderStyle.FixedSingle}
            txtToLoc = New TextBox With {.Dock = DockStyle.Fill, .Text = "", .TabIndex = 2, .BackColor = CivicCalmTheme.ColorSurface, .ForeColor = CivicCalmTheme.ColorInk, .BorderStyle = BorderStyle.FixedSingle}
            txtReason = New TextBox With {.Dock = DockStyle.Fill, .Multiline = True, .Text = "", .TabIndex = 3, .BackColor = CivicCalmTheme.ColorSurface, .ForeColor = CivicCalmTheme.ColorInk, .BorderStyle = BorderStyle.FixedSingle}

            Dim flwButtons As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.RightToLeft,
                .WrapContents = False
            }

            btnCancel = New Button With {
                .Text = " &Cancel",
                .Size = New Size(95, 34),
                .BackColor = CivicCalmTheme.ColorWell,
                .ForeColor = CivicCalmTheme.ColorInk,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(8, 6, 0, 0),
                .Image = AppAssets.GetIcon("x", 14, CivicCalmTheme.ColorInk),
                .ImageAlign = ContentAlignment.MiddleLeft,
                .TextAlign = ContentAlignment.MiddleCenter,
                .TextImageRelation = TextImageRelation.ImageBeforeText,
                .Padding = New Padding(6, 0, 6, 0)
            }
            btnCancel.FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder
            AddHandler btnCancel.Click, Sub()
                                            Me.DialogResult = DialogResult.Cancel
                                            Me.Close()
                                        End Sub

            btnSave = New Button With {
                .Text = " &Transfer Storage",
                .Size = New Size(165, 34),
                .BackColor = CivicCalmTheme.ColorPrimary,
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(0, 6, 0, 0),
                .Image = AppAssets.GetIcon("archive", 16, Color.White),
                .ImageAlign = ContentAlignment.MiddleLeft,
                .TextAlign = ContentAlignment.MiddleCenter,
                .TextImageRelation = TextImageRelation.ImageBeforeText,
                .Padding = New Padding(6, 0, 6, 0)
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
            Dim toLoc = txtToLoc.Text.Trim()
            If String.IsNullOrWhiteSpace(toLoc) Then
                MessageBox.Show("Please enter the new landmark storage location.", "Validation Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtToLoc.Focus()
                Return
            End If

            Dim currentUserId As Integer = If(StaffUserId > 0, StaffUserId, 1)
            If Program.Coordinator IsNot Nothing Then
                Program.Coordinator.MoveStorage(DocID, txtFromLoc.Text.Trim(), toLoc, txtReason.Text.Trim(), StaffName, currentUserId)
            Else
                EmbeddedDB.AddMovementLog(DocID, txtFromLoc.Text.Trim(), toLoc, StaffName, txtReason.Text.Trim())
                EmbeddedDB.LogAudit(StaffName, $"Transferred Doc #{DocID} storage location from {txtFromLoc.Text.Trim()} to {toLoc}")
            End If
            Me.DialogResult = DialogResult.OK
            Me.Close()
        End Sub
    End Class
End Namespace
