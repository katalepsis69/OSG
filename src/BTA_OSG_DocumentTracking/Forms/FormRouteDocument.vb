Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Windows.Forms

Namespace BTA_OSG
    Public Class FormRouteDocument
        Inherits Form

        Private DocID As Integer
        Private StaffName As String
        Private StaffUserId As Integer
        ' Bulk mode: when set, one dialog routes every listed document with the same
        ' To office, action, and remarks; each From office is the document's own holding
        ' section. Nothing (single-document mode) keeps the editable From field.
        Private BulkDocIds As List(Of Integer)

        Private txtFrom As TextBox
        Private cmbTo As ComboBox
        Private cmbAction As ComboBox
        Private txtRemarks As TextBox

        Private btnSave As Button
        Private btnCancel As Button

        ''' <param name="officeHistory">Offices already on file, offered as type-to-filter autocomplete.</param>
        ''' <param name="lastToOffice">The document's most recent routing destination, prefilled ("repeat last route").</param>
        ''' <param name="bulkDocIds">Optional list of document ids to route in one pass.</param>
        Public Sub New(docId As Integer, currentOffice As String, user As String, Optional staffUserId As Integer = 1,
                       Optional officeHistory As List(Of String) = Nothing, Optional lastToOffice As String = "", Optional bulkDocIds As List(Of Integer) = Nothing)
            Me.DocID = docId
            Me.StaffName = user
            Me.StaffUserId = staffUserId
            Me.BulkDocIds = bulkDocIds
            AppAssets.ApplyFormIcon(Me)

            Me.Text = If(bulkDocIds IsNot Nothing,
                String.Format("Log Routing Step for {0} Documents", bulkDocIds.Count),
                "Log Document Office Routing Step")
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

            Dim isBulk = bulkDocIds IsNot Nothing
            txtFrom = New TextBox With {.Dock = DockStyle.Fill, .Text = currentOffice, .TabIndex = 1, .BackColor = CivicCalmTheme.ColorSurface, .ForeColor = CivicCalmTheme.ColorInk, .BorderStyle = BorderStyle.FixedSingle, .ReadOnly = isBulk}
            If isBulk Then
                ' Each document's own current section is the honest From; a single shared
                ' value would record wrong custody legs, so bulk mode states the rule.
                txtFrom.Text = "(each document's current section desk)"
            End If

            ' Editable dropdown seeded with offices already on file: pick or type, free
            ' text still accepted. Prefilled with the document's last hop so routing to
            ' the same desk again needs no retyping.
            cmbTo = New ComboBox With {
                .Dock = DockStyle.Fill,
                .DropDownStyle = ComboBoxStyle.DropDown,
                .TabIndex = 2,
                .DropDownWidth = 460,
                .BackColor = CivicCalmTheme.ColorSurface,
                .ForeColor = CivicCalmTheme.ColorInk,
                .Text = lastToOffice
            }
            Try
                If System.Threading.Thread.CurrentThread.GetApartmentState() = System.Threading.ApartmentState.STA Then
                    cmbTo.AutoCompleteMode = AutoCompleteMode.SuggestAppend
                    cmbTo.AutoCompleteSource = AutoCompleteSource.ListItems
                End If
            Catch ex As System.Threading.ThreadStateException
                ' Suppress in MTA test runners where OLE COM autocomplete is unsupported
            End Try
            If officeHistory IsNot Nothing Then
                For Each officeName As String In officeHistory
                    cmbTo.Items.Add(officeName)
                Next
            End If

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

            btnCancel = CivicCalmTheme.CreateIconButton("&Cancel", "x", 14, margin:=New Padding(8, 6, 0, 0))
            AddHandler btnCancel.Click, Sub()
                                            Me.DialogResult = DialogResult.Cancel
                                            Me.Close()
                                        End Sub

            btnSave = CivicCalmTheme.CreateIconButton("&Route Document", "arrows-split", backColor:=CivicCalmTheme.ColorPrimary, margin:=New Padding(0, 6, 0, 0))
            AddHandler btnSave.Click, AddressOf OnSave

            flwButtons.Controls.AddRange(New Control() {btnCancel, btnSave})

            tblLayout.Controls.Add(New Label With {.Text = "From Office:", .Font = CivicCalmTheme.FontFieldLabel, .ForeColor = CivicCalmTheme.ColorInkMuted, .Dock = DockStyle.Fill, .TextAlign = ContentAlignment.MiddleLeft}, 0, 0)
            tblLayout.Controls.Add(txtFrom, 1, 0)
            tblLayout.Controls.Add(New Label With {.Text = "To Office:", .Font = CivicCalmTheme.FontFieldLabel, .ForeColor = CivicCalmTheme.ColorInkMuted, .Dock = DockStyle.Fill, .TextAlign = ContentAlignment.MiddleLeft}, 0, 1)
            tblLayout.Controls.Add(cmbTo, 1, 1)
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
            Dim toOffice = cmbTo.Text.Trim()
            If String.IsNullOrWhiteSpace(toOffice) Then
                MessageBox.Show("Please enter the destination office (To Office).", "Validation Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbTo.Focus()
                Return
            End If

            ' Same guard as the directive button: a queued second click of a double-click
            ' re-validates fine and would file the transmittal twice.
            btnSave.Enabled = False
            Try
                Dim currentUserId As Integer = If(StaffUserId > 0, StaffUserId, 1)
                If BulkDocIds IsNot Nothing Then
                    SaveBulk(toOffice, currentUserId)
                Else
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
                End If
                Me.DialogResult = DialogResult.OK
                Me.Close()
            Finally
                btnSave.Enabled = True
            End Try
        End Sub

        ''' <summary>
        ''' Bulk routing: each document moves from its own current holding section (the
        ''' destination office when unassigned), with the one shared destination, action,
        ''' and remarks. Every document carries its own guard and error report, so one
        ''' state conflict among many never blocks the rest.
        ''' </summary>
        Private Sub SaveBulk(toOffice As String, currentUserId As Integer)
            Dim routed As Integer = 0
            Dim failures As New List(Of String)()
            For Each id In BulkDocIds
                Dim doc = EmbeddedDB.GetDocumentByID(id)
                If doc Is Nothing Then
                    failures.Add(String.Format("Document {0}: no longer available", id))
                    Continue For
                End If
                Dim fromOffice = If(Not String.IsNullOrWhiteSpace(doc.AssignedSection), doc.AssignedSection, doc.DestinationOffice)
                If String.IsNullOrWhiteSpace(fromOffice) Then fromOffice = "(unassigned)"
                Try
                    If Program.Coordinator IsNot Nothing Then
                        Program.Coordinator.RouteDocument(id, fromOffice, toOffice, cmbAction.Text, txtRemarks.Text.Trim(), StaffName, currentUserId)
                    Else
                        EmbeddedDB.AddRoutingLog(id, fromOffice, toOffice, StaffName, cmbAction.Text, txtRemarks.Text.Trim())
                        EmbeddedDB.LogAudit(StaffName, $"Routed Doc #{id} from {fromOffice} to {toOffice}", actionType:="ROUTE_LOGGED")
                    End If
                    routed += 1
                Catch ex As Exception
                    failures.Add(String.Format("Document {0}: {1}", id, ex.Message))
                End Try
            Next

            If failures.Count > 0 Then
                Dim shown As New List(Of String)(failures)
                If shown.Count > 6 Then
                    shown = shown.GetRange(0, 6)
                    shown.Add(String.Format("...and {0} more.", failures.Count - 6))
                End If
                MessageBox.Show(String.Format("Routed {0} of {1} document(s)." & vbCrLf & vbCrLf & "{2}", routed, BulkDocIds.Count, String.Join(vbCrLf, shown)),
                                "Bulk Routing Partially Completed", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        End Sub
    End Class
End Namespace
