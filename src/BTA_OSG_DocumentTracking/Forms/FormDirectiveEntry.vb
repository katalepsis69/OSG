Imports System.Drawing
Imports System.Windows.Forms


Namespace BTA_OSG
    Public Class FormDirectiveEntry
        Inherits Form

        Private _document As Document
        
        Private lblDocCode As Label
        Private cmbDirectiveType As ComboBox
        Private txtText As TextBox
        Private txtRemarks As TextBox
        Private btnIssue As Button
        Private btnCancel As Button

        Public Sub New(doc As Document)
            _document = doc
            InitializeComponent()
            ApplyTheme()
            LoadTypes()
        End Sub

        Private Sub InitializeComponent()
            Me.Size = New Size(500, 400)
            Me.StartPosition = FormStartPosition.CenterParent
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.Text = "Issue SG Directive"

            Dim pnl = New Panel() With {.Dock = DockStyle.Fill, .Padding = New Padding(10)}
            
            lblDocCode = New Label() With {.Text = "Document: " & _document.DocCode, .Dock = DockStyle.Top, .Height = 30, .Font = New Font("Segoe UI", 10, FontStyle.Bold)}
            
            Dim lblType = New Label() With {.Text = "Directive Type:", .Dock = DockStyle.Top, .Height = 25}
            cmbDirectiveType = New ComboBox() With {.Dock = DockStyle.Top, .DropDownStyle = ComboBoxStyle.DropDownList}
            
            Dim lblText = New Label() With {.Text = "Directive Details:", .Dock = DockStyle.Top, .Height = 25}
            txtText = New TextBox() With {.Dock = DockStyle.Top, .Multiline = True, .Height = 80}
            
            Dim lblRemarks = New Label() With {.Text = "Remarks:", .Dock = DockStyle.Top, .Height = 25}
            txtRemarks = New TextBox() With {.Dock = DockStyle.Top, .Multiline = True, .Height = 60}
            
            btnIssue = New Button() With {.Text = "Issue Directive", .Dock = DockStyle.Top, .Height = 40, .FlatStyle = FlatStyle.Flat}
            AddHandler btnIssue.Click, AddressOf IssueBtn_Click
            
            btnCancel = New Button() With {.Text = "Cancel", .Dock = DockStyle.Top, .Height = 40, .FlatStyle = FlatStyle.Flat}
            AddHandler btnCancel.Click, Sub(s, e) Me.DialogResult = DialogResult.Cancel

            pnl.Controls.Add(btnCancel)
            pnl.Controls.Add(btnIssue)
            pnl.Controls.Add(txtRemarks)
            pnl.Controls.Add(lblRemarks)
            pnl.Controls.Add(txtText)
            pnl.Controls.Add(lblText)
            pnl.Controls.Add(cmbDirectiveType)
            pnl.Controls.Add(lblType)
            pnl.Controls.Add(lblDocCode)
            
            pnl.Controls.SetChildIndex(lblDocCode, 0)
            pnl.Controls.SetChildIndex(lblType, 1)
            pnl.Controls.SetChildIndex(cmbDirectiveType, 2)
            pnl.Controls.SetChildIndex(lblText, 3)
            pnl.Controls.SetChildIndex(txtText, 4)
            pnl.Controls.SetChildIndex(lblRemarks, 5)
            pnl.Controls.SetChildIndex(txtRemarks, 6)
            pnl.Controls.SetChildIndex(btnIssue, 7)
            pnl.Controls.SetChildIndex(btnCancel, 8)

            Me.Controls.Add(pnl)
        End Sub

        Private Sub ApplyTheme()
            Me.BackColor = Color.FromArgb(15, 23, 42)
            For Each c As Control In Me.Controls(0).Controls
                If TypeOf c Is Label Then c.ForeColor = Color.White
            Next
            btnIssue.BackColor = Color.FromArgb(79, 70, 229)
            btnIssue.ForeColor = Color.White
            btnCancel.BackColor = Color.FromArgb(30, 41, 59)
            btnCancel.ForeColor = Color.White
        End Sub

        Private Sub LoadTypes()
            ' Mocked load from ReferenceDataRepo
            cmbDirectiveType.Items.Add("APPROVAL")
            cmbDirectiveType.Items.Add("REVIEW")
            cmbDirectiveType.Items.Add("FOR_INFO")
            If cmbDirectiveType.Items.Count > 0 Then cmbDirectiveType.SelectedIndex = 0
        End Sub

        Private Sub IssueBtn_Click(sender As Object, e As EventArgs)
            Try
                Dim d = New ActionDirective() With {
                    .DocumentID = _document.DocumentID,
                    .DirectiveTypeID = cmbDirectiveType.SelectedIndex + 1,
                    .DirectiveText = txtText.Text,
                    .Remarks = txtRemarks.Text,
                    .IssuedByUserID = SessionManager.CurrentSession.User.UserID,
                    .IssuedAtUTC = Date.UtcNow,
                    .IsActive = True
                }
                ' AppStartup.DirectiveService.IssueDirective(d) ' Placeholder actual call
                MessageBox.Show("Directive issued successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.DialogResult = DialogResult.OK
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub
    End Class
End Namespace
