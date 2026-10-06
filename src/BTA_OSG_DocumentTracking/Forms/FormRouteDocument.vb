Option Explicit On
Option Strict On

Imports System
Imports System.Drawing
Imports System.Windows.Forms

Namespace BTA_OSG
    Public Class FormRouteDocument
        Inherits Form

        Private DocID As Integer
        Private StaffName As String
        Private StaffUserId As Integer
        Private txtFrom As TextBox
        Private txtTo As TextBox
        Private cmbAction As ComboBox
        Private txtRemarks As TextBox

        Private btnSave As Button
        Private btnCancel As Button

        Public Sub New(docId As Integer, currentOffice As String, user As String, Optional staffUserId As Integer = 1)
            Me.DocID = docId
            Me.StaffName = user
            Me.StaffUserId = staffUserId
            AppAssets.ApplyFormIcon(Me)

            Me.Text = "Log Document Office Routing Step"
            Me.Size = New Size(500, 360)
            Me.MinimumSize = New Size(460, 340)
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
                .RowCount = 5
            }
            tblLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 120.0F))
            tblLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            tblLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 40.0F))
            tblLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 40.0F))
            tblLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 40.0F))
            tblLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 60.0F))
            tblLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 50.0F))

            txtFrom = New TextBox With {.Dock = DockStyle.Fill, .Text = currentOffice, .TabIndex = 1, .BackColor = CivicCalmTheme.ColorSurface, .ForeColor = CivicCalmTheme.ColorInk, .BorderStyle = BorderStyle.FixedSingle}
            txtTo = New TextBox With {.Dock = DockStyle.Fill, .Text = "", .TabIndex = 2, .BackColor = CivicCalmTheme.ColorSurface, .ForeColor = CivicCalmTheme.ColorInk, .BorderStyle = BorderStyle.FixedSingle}
            ' A free-text action used to default to FOR_TRANSMITTAL, which is not a status
            ' code, and the connected path then regressed the document to RECEIVED. The
            ' dropdown carries the seeded status codes only; ROUTED is the neutral default.
            cmbAction = New ComboBox With {.Dock = DockStyle.Fill, .DropDownStyle = ComboBoxStyle.DropDownList, .TabIndex = 3, .BackColor = CivicCalmTheme.ColorSurface, .ForeColor = CivicCalmTheme.ColorInk, .FlatStyle = FlatStyle.Flat}
            cmbAction.Items.AddRange(DocumentStatus.RouteTargetStatusCodes)
            cmbAction.SelectedIndex = 0
            txtRemarks = New TextBox With {.Dock = DockStyle.Fill, .Multiline = True, .Text = "", .TabIndex = 4, .BackColor = CivicCalmTheme.ColorSurface, .ForeColor = CivicCalmTheme.ColorInk, .BorderStyle = BorderStyle.FixedSingle}

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
                .Text = " &Route Document",
                .Size = New Size(155, 34),
                .BackColor = CivicCalmTheme.ColorPrimary,
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(0, 6, 0, 0),
                .Image = AppAssets.GetIcon("arrows-split", 16, Color.White),
                .ImageAlign = ContentAlignment.MiddleLeft,
                .TextAlign = ContentAlignment.MiddleCenter,
                .TextImageRelation = TextImageRelation.ImageBeforeText,
                .Padding = New Padding(6, 0, 6, 0)
            }
            btnSave.FlatAppearance.BorderSize = 0
            AddHandler btnSave.Click, AddressOf OnSave

            flwButtons.Controls.AddRange(New Control() {btnCancel, btnSave})

            tblLayout.Controls.Add(New Label With {.Text = "From Office:", .Font = CivicCalmTheme.FontFieldLabel, .ForeColor = CivicCalmTheme.ColorInkMuted, .Dock = DockStyle.Fill, .TextAlign = ContentAlignment.MiddleLeft}, 0, 0)
            tblLayout.Controls.Add(txtFrom, 1, 0)
            tblLayout.Controls.Add(New Label With {.Text = "To Office:", .Font = CivicCalmTheme.FontFieldLabel, .ForeColor = CivicCalmTheme.ColorInkMuted, .Dock = DockStyle.Fill, .TextAlign = ContentAlignment.MiddleLeft}, 0, 1)
            tblLayout.Controls.Add(txtTo, 1, 1)
            tblLayout.Controls.Add(New Label With {.Text = "Action Taken:", .Font = CivicCalmTheme.FontFieldLabel, .ForeColor = CivicCalmTheme.ColorInkMuted, .Dock = DockStyle.Fill, .TextAlign = ContentAlignment.MiddleLeft}, 0, 2)
            tblLayout.Controls.Add(cmbAction, 1, 2)
            tblLayout.Controls.Add(New Label With {.Text = "Remarks:", .Font = CivicCalmTheme.FontFieldLabel, .ForeColor = CivicCalmTheme.ColorInkMuted, .Dock = DockStyle.Fill, .TextAlign = ContentAlignment.TopLeft}, 0, 3)
            tblLayout.Controls.Add(txtRemarks, 1, 3)
            tblLayout.Controls.Add(flwButtons, 1, 4)

            pnlCard.Controls.Add(tblLayout)
            Me.Controls.Add(pnlCard)

            Me.AcceptButton = btnSave
            Me.CancelButton = btnCancel
        End Sub

        Private Sub OnSave(sender As Object, e As EventArgs)
            Dim toOffice = txtTo.Text.Trim()
            If String.IsNullOrWhiteSpace(toOffice) Then
                MessageBox.Show("Please enter the destination office (To Office).", "Validation Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtTo.Focus()
                Return
            End If

            ' Same guard as the directive button: a queued second click of a double-click
            ' re-validates fine and would file the transmittal twice.
            btnSave.Enabled = False
            Try
                Dim currentUserId As Integer = If(StaffUserId > 0, StaffUserId, 1)
                Try
                    If Program.Coordinator IsNot Nothing Then
                        Program.Coordinator.RouteDocument(DocID, txtFrom.Text.Trim(), toOffice, cmbAction.Text, txtRemarks.Text.Trim(), StaffName, currentUserId)
                    Else
                        EmbeddedDB.AddRoutingLog(DocID, txtFrom.Text.Trim(), toOffice, StaffName, cmbAction.Text, txtRemarks.Text.Trim())
                        EmbeddedDB.LogAudit(StaffName, $"Routed Doc #{DocID} from {txtFrom.Text.Trim()} to {toOffice}", actionType:="ROUTE_LOGGED")
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message, "Route Not Logged", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End Try
                Me.DialogResult = DialogResult.OK
                Me.Close()
            Finally
                btnSave.Enabled = True
            End Try
        End Sub
    End Class
End Namespace
