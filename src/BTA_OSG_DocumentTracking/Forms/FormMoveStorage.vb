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

        Public Sub New(docId As Integer, currentLoc As String, user As String)
            Me.DocID = docId
            Me.StaffName = user

            Me.Text = "Transfer Physical Landmark Storage Location"
            Me.Size = New Size(470, 300)
            Me.StartPosition = FormStartPosition.CenterParent
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.BackColor = Color.FromArgb(15, 23, 42)

            txtFromLoc = New TextBox With {.Location = New Point(150, 20), .Width = 270, .Text = currentLoc, .ReadOnly = True, .BackColor = Color.FromArgb(30, 41, 59), .ForeColor = Color.White, .BorderStyle = BorderStyle.FixedSingle}
            txtToLoc = New TextBox With {.Location = New Point(150, 60), .Width = 270, .Text = "CAB-A/S-3/BOX-05", .BackColor = Color.FromArgb(30, 41, 59), .ForeColor = Color.White, .BorderStyle = BorderStyle.FixedSingle}
            txtReason = New TextBox With {.Location = New Point(150, 100), .Width = 270, .Text = "Archival reorganization per OSG directive.", .BackColor = Color.FromArgb(30, 41, 59), .ForeColor = Color.White, .BorderStyle = BorderStyle.FixedSingle}

            Dim btnSave As New Button With {.Text = "Save Physical Transfer", .Location = New Point(150, 155), .Size = New Size(180, 36), .BackColor = Color.FromArgb(14, 165, 233), .ForeColor = Color.White, .FlatStyle = FlatStyle.Flat, .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold)}
            btnSave.FlatAppearance.BorderSize = 0
            AddHandler btnSave.Click, AddressOf OnSave

            Me.Controls.AddRange(New Control() {
                New Label With {.Text = "Current Location:", .Location = New Point(20, 23), .AutoSize = True, .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold), .ForeColor = Color.FromArgb(148, 163, 184)}, txtFromLoc,
                New Label With {.Text = "New (Cab/Shelf/Box):", .Location = New Point(20, 63), .AutoSize = True, .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold), .ForeColor = Color.FromArgb(148, 163, 184)}, txtToLoc,
                New Label With {.Text = "Reason for Move:", .Location = New Point(20, 103), .AutoSize = True, .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold), .ForeColor = Color.FromArgb(148, 163, 184)}, txtReason,
                btnSave
            })
        End Sub

        Private Sub OnSave(sender As Object, e As EventArgs)
            EmbeddedDB.AddMovementLog(DocID, txtFromLoc.Text.Trim(), txtToLoc.Text.Trim(), StaffName, txtReason.Text.Trim())
            EmbeddedDB.LogAudit(StaffName, String.Format("Transferred Doc #{0} physical location from {1} to {2}", DocID, txtFromLoc.Text, txtToLoc.Text))
            Me.DialogResult = DialogResult.OK
            Me.Close()
        End Sub
    End Class
End Namespace
