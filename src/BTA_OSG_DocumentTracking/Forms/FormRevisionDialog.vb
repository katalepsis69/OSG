Option Explicit On
Option Strict On

Imports System
Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms

Namespace BTA_OSG
    Public Class FormRevisionDialog
        Inherits Form

        <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
        Public Property PunchlistNotes As String = ""

        <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
        Public Property TargetSection As String = ""

        Private txtPunchlist As TextBox
        Private cmbSection As ComboBox
        Private btnSubmit As Button
        Private btnCancel As Button

        Public Sub New(docId As Integer, docCode As String, docTitle As String, defaultSection As String)
            AppAssets.ApplyFormIcon(Me)
            Me.Text = String.Format("Sec Gen Revision Request : [{0}]", docCode)
            Me.Size = New Size(560, 440)
            Me.MinimumSize = New Size(480, 380)
            Me.StartPosition = FormStartPosition.CenterParent
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.BackColor = CivicCalmTheme.ColorCanvas
            Me.Font = CivicCalmTheme.FontBody

            Dim pnlCard As New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(20)
            }

            Dim lblHeader As New Label With {
                .Text = "Secretary-General Document Revision Order",
                .Font = CivicCalmTheme.FontSectionHeader,
                .ForeColor = CivicCalmTheme.ColorInk,
                .Dock = DockStyle.Top,
                .Height = 28
            }

            Dim lblDocDesc As New Label With {
                .Text = String.Format("Document #{0} [{1}] : {2}", docId, docCode, docTitle),
                .Font = CivicCalmTheme.FontFieldLabel,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .Dock = DockStyle.Top,
                .Height = 24
            }

            Dim tblLayout As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 2,
                .RowCount = 3
            }
            tblLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 150.0F))
            tblLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            tblLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 42.0F))
            tblLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
            tblLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 48.0F))

            Dim lblSec As New Label With {
                .Text = "Return to Section:",
                .Font = CivicCalmTheme.FontFieldLabel,
                .ForeColor = CivicCalmTheme.ColorInk,
                .Dock = DockStyle.Fill,
                .TextAlign = ContentAlignment.MiddleLeft
            }
            cmbSection = New ComboBox With {
                .Dock = DockStyle.Fill,
                .DropDownStyle = ComboBoxStyle.DropDownList,
                .Font = CivicCalmTheme.FontBody,
                .BackColor = CivicCalmTheme.ColorSurface,
                .ForeColor = CivicCalmTheme.ColorInk
            }
            cmbSection.Items.AddRange(New Object() {"Secretariat", "Legislative Section", "Finance Section", "Travel Section", "Records Section"})
            Dim secIdx = cmbSection.Items.IndexOf(defaultSection)
            cmbSection.SelectedIndex = If(secIdx >= 0, secIdx, 0)

            Dim lblPunch As New Label With {
                .Text = "Punchlist Instructions:" & vbCrLf & "(Items to amend)",
                .Font = CivicCalmTheme.FontFieldLabel,
                .ForeColor = CivicCalmTheme.ColorInk,
                .Dock = DockStyle.Fill,
                .TextAlign = ContentAlignment.TopLeft,
                .Padding = New Padding(0, 6, 0, 0)
            }
            txtPunchlist = New TextBox With {
                .Dock = DockStyle.Fill,
                .Multiline = True,
                .ScrollBars = ScrollBars.Vertical,
                .Font = CivicCalmTheme.FontBody,
                .BackColor = CivicCalmTheme.ColorSurface,
                .ForeColor = CivicCalmTheme.ColorInk,
                .BorderStyle = BorderStyle.FixedSingle
            }

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
            AddHandler btnCancel.Click, Sub() Me.DialogResult = DialogResult.Cancel

            btnSubmit = New Button With {
                .Text = " &Issue Revision Order",
                .Size = New Size(185, 34),
                .BackColor = CivicCalmTheme.ColorDanger,
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(0, 6, 0, 0),
                .Image = AppAssets.GetIcon("arrow-clockwise", 16, Color.White),
                .ImageAlign = ContentAlignment.MiddleLeft,
                .TextAlign = ContentAlignment.MiddleCenter,
                .TextImageRelation = TextImageRelation.ImageBeforeText,
                .Padding = New Padding(6, 0, 6, 0)
            }
            btnSubmit.FlatAppearance.BorderSize = 0
            AddHandler btnSubmit.Click, AddressOf OnSubmit

            flwButtons.Controls.AddRange(New Control() {btnCancel, btnSubmit})

            tblLayout.Controls.Add(lblSec, 0, 0)
            tblLayout.Controls.Add(cmbSection, 1, 0)
            tblLayout.Controls.Add(lblPunch, 0, 1)
            tblLayout.Controls.Add(txtPunchlist, 1, 1)
            tblLayout.Controls.Add(flwButtons, 1, 2)

            pnlCard.Controls.Add(tblLayout)
            pnlCard.Controls.Add(lblDocDesc)
            pnlCard.Controls.Add(lblHeader)

            Me.Controls.Add(pnlCard)
            Me.AcceptButton = btnSubmit
            Me.CancelButton = btnCancel
        End Sub

        Private Sub OnSubmit(sender As Object, e As EventArgs)
            Dim notes = txtPunchlist.Text.Trim()
            If String.IsNullOrWhiteSpace(notes) Then
                MessageBox.Show("Please enter the specific revision punchlist instructions before submitting.", "Validation Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtPunchlist.Focus()
                Return
            End If

            PunchlistNotes = notes
            TargetSection = If(cmbSection.SelectedItem IsNot Nothing, cmbSection.SelectedItem.ToString(), "Records Section")
            Me.DialogResult = DialogResult.OK
            Me.Close()
        End Sub
    End Class
End Namespace
