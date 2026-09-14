Imports System.Drawing
Imports System.Windows.Forms

Namespace BTA_OSG
    Public Class FormRouteDocument
        Inherits Form

        Private DocID As Integer
        Private StaffName As String
        Private txtFrom As TextBox
        Private txtTo As TextBox
        Private txtAction As TextBox
        Private txtRemarks As TextBox

        Public Sub New(docId As Integer, currentOffice As String, user As String)
            Me.DocID = docId
            Me.StaffName = user

            Me.Text = "Log Document Office Routing Step"
            Me.Size = New Size(460, 340)
            Me.StartPosition = FormStartPosition.CenterParent
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.BackColor = Color.FromArgb(15, 23, 42)

            txtFrom = New TextBox With {.Location = New Point(140, 20), .Width = 270, .Text = currentOffice, .BackColor = Color.FromArgb(30, 41, 59), .ForeColor = Color.White, .BorderStyle = BorderStyle.FixedSingle}
            txtTo = New TextBox With {.Location = New Point(140, 60), .Width = 270, .Text = "Speaker's Office", .BackColor = Color.FromArgb(30, 41, 59), .ForeColor = Color.White, .BorderStyle = BorderStyle.FixedSingle}
            txtAction = New TextBox With {.Location = New Point(140, 100), .Width = 270, .Text = "FOR_SIGNATURE", .BackColor = Color.FromArgb(30, 41, 59), .ForeColor = Color.White, .BorderStyle = BorderStyle.FixedSingle}
            txtRemarks = New TextBox With {.Location = New Point(140, 140), .Width = 270, .Text = "Transmitted for Speaker approval.", .BackColor = Color.FromArgb(30, 41, 59), .ForeColor = Color.White, .BorderStyle = BorderStyle.FixedSingle}

            Dim btnSave As New Button With {.Text = "Save Route Step", .Location = New Point(140, 195), .Size = New Size(160, 36), .BackColor = Color.FromArgb(79, 70, 229), .ForeColor = Color.White, .FlatStyle = FlatStyle.Flat, .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold)}
            btnSave.FlatAppearance.BorderSize = 0
            AddHandler btnSave.Click, AddressOf OnSave

            Me.Controls.AddRange(New Control() {
                New Label With {.Text = "From Office:", .Location = New Point(20, 23), .AutoSize = True, .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold), .ForeColor = Color.FromArgb(148, 163, 184)}, txtFrom,
                New Label With {.Text = "To Office:", .Location = New Point(20, 63), .AutoSize = True, .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold), .ForeColor = Color.FromArgb(148, 163, 184)}, txtTo,
                New Label With {.Text = "Action Taken:", .Location = New Point(20, 103), .AutoSize = True, .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold), .ForeColor = Color.FromArgb(148, 163, 184)}, txtAction,
                New Label With {.Text = "Remarks:", .Location = New Point(20, 143), .AutoSize = True, .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold), .ForeColor = Color.FromArgb(148, 163, 184)}, txtRemarks,
                btnSave
            })
        End Sub

        Private Sub OnSave(sender As Object, e As EventArgs)
            EmbeddedDB.AddRoutingLog(DocID, txtFrom.Text.Trim(), txtTo.Text.Trim(), StaffName, txtAction.Text.Trim(), txtRemarks.Text.Trim())
            EmbeddedDB.LogAudit(StaffName, String.Format("Routed Doc #{0} from {1} to {2}", DocID, txtFrom.Text, txtTo.Text))
            Me.DialogResult = DialogResult.OK
            Me.Close()
        End Sub
    End Class
End Namespace
