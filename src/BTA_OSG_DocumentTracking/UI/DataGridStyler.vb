Option Explicit On
Option Strict On

Imports System.Drawing
Imports System.Reflection
Imports System.Windows.Forms

Namespace BTA_OSG
    Public NotInheritable Class DataGridStyler
        Private Sub New()
        End Sub

        ' Badge fills need a gutter so a status colour never touches the neighbouring column's text.
        Private Shared ReadOnly CellGutter As New Padding(8, 2, 8, 2)

        ' A Drive URL is one unbreakable token, so rendering it raw clips the text and inflates the
        ' column; the grid states whether a soft copy is attached and keeps the link in the tooltip.
        Private Const SoftCopyLinkLabel As String = "Attached"
        Private Const SoftCopyMissingLabel As String = "Not Attached"

        Public Shared Sub ApplyCivicStyle(dgv As DataGridView)
            If dgv Is Nothing Then Return

            dgv.AutoGenerateColumns = True
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            dgv.EnableHeadersVisualStyles = False
            dgv.BackgroundColor = CivicCalmTheme.ColorSurface
            dgv.BorderStyle = BorderStyle.FixedSingle
            dgv.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
            dgv.GridColor = CivicCalmTheme.ColorBorder
            dgv.ColumnHeadersHeight = 32
            dgv.RowTemplate.Height = 28
            dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            dgv.MultiSelect = False
            dgv.RowHeadersVisible = False

            ' Enable DoubleBuffered via reflection to prevent flicker during fast scrolling
            Dim prop = GetType(Control).GetProperty("DoubleBuffered", BindingFlags.Instance Or BindingFlags.NonPublic)
            If prop IsNot Nothing Then
                prop.SetValue(dgv, True, Nothing)
            End If

            ' Column Headers (Tier 2 Recessed Well)
            dgv.ColumnHeadersDefaultCellStyle.BackColor = CivicCalmTheme.ColorWell
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = CivicCalmTheme.ColorInk
            dgv.ColumnHeadersDefaultCellStyle.Font = CivicCalmTheme.FontSectionHeader

            ' Default Row Styling (Pure White)
            dgv.DefaultCellStyle.BackColor = CivicCalmTheme.ColorSurface
            dgv.DefaultCellStyle.ForeColor = CivicCalmTheme.ColorInk
            dgv.DefaultCellStyle.Font = CivicCalmTheme.FontTabular
            dgv.DefaultCellStyle.SelectionBackColor = CivicCalmTheme.ColorPrimarySoft
            dgv.DefaultCellStyle.SelectionForeColor = CivicCalmTheme.ColorPrimary
            dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True

            ' Alternating Row Styling (Soft Off-White Tint #F9FAFB)
            dgv.AlternatingRowsDefaultCellStyle.BackColor = ColorTranslator.FromHtml("#F9FAFB")
            dgv.AlternatingRowsDefaultCellStyle.ForeColor = CivicCalmTheme.ColorInk
            dgv.AlternatingRowsDefaultCellStyle.SelectionBackColor = CivicCalmTheme.ColorPrimarySoft
            dgv.AlternatingRowsDefaultCellStyle.SelectionForeColor = CivicCalmTheme.ColorPrimary
            dgv.AlternatingRowsDefaultCellStyle.WrapMode = DataGridViewTriState.True

            ' Auto-size row heights to always display full content without truncation.
            ' DisplayedCellsExceptHeaders bounds each layout pass to the rows on screen;
            ' rows re-measure as they scroll into view, so large registries stay responsive.
            dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCellsExceptHeaders

            RemoveHandler dgv.CellFormatting, AddressOf HandleCellFormatting
            AddHandler dgv.CellFormatting, AddressOf HandleCellFormatting
        End Sub

        Private Shared Sub ApplyStatusBadge(e As DataGridViewCellFormattingEventArgs, backColor As Color, foreColor As Color)
            e.CellStyle.BackColor = backColor
            e.CellStyle.ForeColor = foreColor
        End Sub

        Private Shared Sub HandleCellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs)
            Dim dgv = TryCast(sender, DataGridView)
            If dgv Is Nothing OrElse e.RowIndex < 0 OrElse e.ColumnIndex < 0 Then Return

            Select Case dgv.Columns(e.ColumnIndex).Name
                Case "CurrentStatus"
                    FormatStatusCell(e)
                Case "GDriveURL"
                    FormatSoftCopyCell(dgv, e)
            End Select
        End Sub

        Private Shared Sub FormatStatusCell(e As DataGridViewCellFormattingEventArgs)
            Dim code As String = ""
            If e.Value IsNot Nothing AndAlso Not Convert.IsDBNull(e.Value) Then
                code = e.Value.ToString().Trim().ToUpperInvariant()
            End If
            If code = "" Then Return

            If code = "FOR_REVISION" OrElse code.Contains("REVISION") Then
                ApplyStatusBadge(e, CivicCalmTheme.ColorStatusRevisionBg, CivicCalmTheme.ColorStatusRevisionFg)
            ElseIf code = "FOR_REVIEW" OrElse code.Contains("REVIEW") Then
                ApplyStatusBadge(e, CivicCalmTheme.ColorStatusReviewBg, CivicCalmTheme.ColorStatusReviewFg)
            ElseIf code = "APPROVED" Then
                ApplyStatusBadge(e, CivicCalmTheme.ColorStatusApprovedBg, CivicCalmTheme.ColorStatusApprovedFg)
            ElseIf code = "RELEASED" Then
                ApplyStatusBadge(e, CivicCalmTheme.ColorStatusReleasedBg, CivicCalmTheme.ColorStatusReleasedFg)
            ElseIf code = "RECEIVED" OrElse code = "LOGGED" Then
                ApplyStatusBadge(e, CivicCalmTheme.ColorStatusReceivedBg, CivicCalmTheme.ColorStatusReceivedFg)
            ElseIf code = "PENDING" OrElse code = "IN_TRANSIT" OrElse code = "IN TRANSIT" Then
                ApplyStatusBadge(e, CivicCalmTheme.ColorStatusPendingBg, CivicCalmTheme.ColorStatusPendingFg)
            ElseIf code = "ARCHIVED" OrElse code = "FILED" Then
                ApplyStatusBadge(e, CivicCalmTheme.ColorStatusArchivedBg, CivicCalmTheme.ColorStatusArchivedFg)
            Else
                ' Unrecognized status text stays exactly as stored rather than being rewritten.
                Return
            End If

            e.Value = DocumentStatus.DisplayName(code)
            e.FormattingApplied = True
        End Sub

        Private Shared Sub FormatSoftCopyCell(dgv As DataGridView, e As DataGridViewCellFormattingEventArgs)
            Dim url As String = ""
            If e.Value IsNot Nothing AndAlso Not Convert.IsDBNull(e.Value) Then
                url = e.Value.ToString().Trim()
            End If

            If url = "" Then
                e.Value = SoftCopyMissingLabel
                e.CellStyle.ForeColor = CivicCalmTheme.ColorInkMuted
                e.CellStyle.Font = CivicCalmTheme.FontMicrocopy
            Else
                e.Value = SoftCopyLinkLabel
                e.CellStyle.ForeColor = CivicCalmTheme.ColorPrimary
                e.CellStyle.Font = CivicCalmTheme.FontTabular

                Dim cell = dgv.Rows(e.RowIndex).Cells(e.ColumnIndex)
                If cell.ToolTipText <> url Then cell.ToolTipText = url
            End If

            e.FormattingApplied = True
        End Sub

        Public Shared Sub SetEmptyState(dgv As DataGridView, watermarkLabel As Label, Optional customMessage As String = Nothing)
            If dgv IsNot Nothing Then
                dgv.Visible = True
            End If
            If watermarkLabel IsNot Nothing Then
                watermarkLabel.Text = If(customMessage, "No documents match the current filter criteria. Press Alt+C to clear filters.")
                watermarkLabel.ForeColor = CivicCalmTheme.ColorInkMuted
                watermarkLabel.Visible = True
                watermarkLabel.BringToFront()
            End If
        End Sub

        Public Shared Sub SetLoadingState(dgv As DataGridView, watermarkLabel As Label)
            If watermarkLabel IsNot Nothing Then
                watermarkLabel.Text = "Loading document records"
                watermarkLabel.ForeColor = CivicCalmTheme.ColorInkMuted
                watermarkLabel.Visible = True
                watermarkLabel.BringToFront()
            End If
        End Sub

        Public Shared Sub SetErrorState(dgv As DataGridView, watermarkLabel As Label, errorMessage As String)
            If watermarkLabel IsNot Nothing Then
                watermarkLabel.Text = "Unable to retrieve records: " & errorMessage
                watermarkLabel.ForeColor = CivicCalmTheme.ColorDanger
                watermarkLabel.Visible = True
                watermarkLabel.BringToFront()
            End If
        End Sub

        Public Shared Sub SetPopulatedState(dgv As DataGridView, watermarkLabel As Label)
            If watermarkLabel IsNot Nothing Then
                watermarkLabel.Visible = False
            End If
            If dgv IsNot Nothing Then
                dgv.Visible = True
            End If
        End Sub

        Private Shared Sub ConfigureColumn(col As DataGridViewColumn, headerText As String, Optional alignment As DataGridViewContentAlignment = DataGridViewContentAlignment.MiddleLeft, Optional font As Font = Nothing)
            col.HeaderText = headerText
            col.DefaultCellStyle.WrapMode = DataGridViewTriState.True
            col.DefaultCellStyle.Alignment = alignment
            col.DefaultCellStyle.Padding = CellGutter
            col.HeaderCell.Style.Alignment = alignment
            If font IsNot Nothing Then
                col.DefaultCellStyle.Font = font
            End If
        End Sub

        ' Columns share the grid width in proportion to the widest text they display (measured as
        ' formatted, never as stored); when that cannot fit, widths are literal and the grid scrolls.
        Private Shared Sub BalanceColumns(dgv As DataGridView)
            Const sampleRows As Integer = 60   ' Bounds measuring overhead during tab refresh
            Const cellPadding As Integer = 18  ' 8px cell gutter each side plus 2px slack
            Dim cellFont = dgv.DefaultCellStyle.Font
            Dim headerFont = dgv.ColumnHeadersDefaultCellStyle.Font

            Dim needs As New List(Of DataGridViewColumn)()
            Dim weights As New List(Of Single)()
            Dim totalNeeded As Single = 0

            For Each col As DataGridViewColumn In dgv.Columns
                If Not col.Visible Then Continue For

                Dim widest As Single = TextRenderer.MeasureText(col.HeaderText, headerFont).Width + cellPadding
                Dim lastRowIndex As Integer = Math.Min(dgv.RowCount - 1, sampleRows)
                For i As Integer = 0 To lastRowIndex
                    Dim cell = dgv.Rows(i).Cells(col.Index)
                    Dim shown = Convert.ToString(cell.FormattedValue)
                    If Not String.IsNullOrEmpty(shown) Then
                        Dim shownFont = If(cell.InheritedStyle.Font, cellFont)
                        Dim needed As Single = TextRenderer.MeasureText(shown, shownFont).Width + cellPadding
                        If needed > widest Then widest = needed
                    End If
                Next

                ' One very long value must not drag every column out of Fill, so a single cell is
                ' allowed to claim at most this much before it is expected to wrap.
                widest = Math.Min(widest, 260.0F)

                needs.Add(col)
                weights.Add(Math.Max(24.0F, widest))
                totalNeeded += Math.Max(24.0F, widest)
            Next

            Dim fits As Boolean = totalNeeded <= dgv.DisplayRectangle.Width
            For i As Integer = 0 To needs.Count - 1
                needs(i).FillWeight = weights(i)
                If fits Then
                    needs(i).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                Else
                    needs(i).AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                    needs(i).Width = CInt(weights(i))
                End If
            Next
        End Sub

        Private Shared Sub FinishLayout(dgv As DataGridView)
            BalanceColumns(dgv)
            dgv.AutoResizeRows(DataGridViewAutoSizeRowsMode.DisplayedCellsExceptHeaders)
        End Sub

        Public Shared Sub FormatDocumentColumns(dgv As DataGridView)
            If dgv Is Nothing OrElse dgv.Columns.Count = 0 Then Return

            Try
                dgv.SuspendLayout()

                For Each col As DataGridViewColumn In dgv.Columns
                    Select Case col.Name
                        Case "DocumentID"
                            ConfigureColumn(col, "ID", DataGridViewContentAlignment.MiddleRight)
                        Case "DocCode"
                            ConfigureColumn(col, "Document Code", font:=CivicCalmTheme.FontIdentifier)
                        Case "DocType"
                            ConfigureColumn(col, "Classification")
                        Case "Title"
                            ConfigureColumn(col, "Document Title / Subject")
                        Case "FlowDirection"
                            ConfigureColumn(col, "Flow", DataGridViewContentAlignment.MiddleCenter)
                        Case "AssignedSection"
                            ConfigureColumn(col, "Assigned Section")
                        Case "OriginatingOffice"
                            ConfigureColumn(col, "Origin Office")
                        Case "DestinationOffice"
                            ConfigureColumn(col, "Destination")
                        Case "CurrentStatus"
                            ' Bold badge font is declared on the column so measurement reserves its width.
                            ConfigureColumn(col, "Status", DataGridViewContentAlignment.MiddleCenter, CivicCalmTheme.FontFieldLabel)
                        Case "TargetDeadlineUTC"
                            ConfigureColumn(col, "Target Deadline")
                        Case "LastActionTaken"
                            ConfigureColumn(col, "Last Action Taken")
                        Case "AssignedStaff"
                            ConfigureColumn(col, "Assigned Staff")
                        Case "DateReceived"
                            ConfigureColumn(col, "Date Received")
                        Case "GDriveURL"
                            ConfigureColumn(col, "Soft Copy Link")
                        Case "CabinetID", "ShelfNo", "BoxCode", "RevisionPunchlist", "ExternalControlNumber"
                            col.Visible = False
                    End Select
                Next
                FinishLayout(dgv)
            Finally
                dgv.ResumeLayout()
            End Try
        End Sub

        Public Shared Sub FormatDirectiveColumns(dgv As DataGridView)
            If dgv Is Nothing OrElse dgv.Columns.Count = 0 Then Return

            Try
                dgv.SuspendLayout()

                For Each col As DataGridViewColumn In dgv.Columns
                    Select Case col.Name
                        Case "DirectiveID"
                            ConfigureColumn(col, "ID", DataGridViewContentAlignment.MiddleRight)
                        Case "DocumentID"
                            ConfigureColumn(col, "Doc ID", DataGridViewContentAlignment.MiddleRight)
                        Case "SGDirective"
                            ConfigureColumn(col, "Directive Action")
                        Case "AssignedTo"
                            ConfigureColumn(col, "Assigned Staff")
                        Case "Notes"
                            ConfigureColumn(col, "Directive Notes / Remarks")
                        Case "LogUser"
                            ConfigureColumn(col, "Issued By")
                        Case "Timestamp"
                            ConfigureColumn(col, "Date / Time Issued")
                    End Select
                Next
                FinishLayout(dgv)
            Finally
                dgv.ResumeLayout()
            End Try
        End Sub

        Public Shared Sub FormatUserColumns(dgv As DataGridView)
            If dgv Is Nothing OrElse dgv.Columns.Count = 0 Then Return

            Try
                dgv.SuspendLayout()

                For Each col As DataGridViewColumn In dgv.Columns
                    Select Case col.Name
                        Case "UserID"
                            ConfigureColumn(col, "User ID", DataGridViewContentAlignment.MiddleRight)
                        Case "RFID_UID"
                            ConfigureColumn(col, "RFID Badge UID", font:=CivicCalmTheme.FontIdentifier)
                        Case "FullName"
                            ConfigureColumn(col, "Full Name")
                        Case "Role"
                            ConfigureColumn(col, "Designated Role")
                        Case "Office"
                            ConfigureColumn(col, "Assigned Section Desk")
                        Case "IsActive"
                            ConfigureColumn(col, "Active", DataGridViewContentAlignment.MiddleCenter)
                    End Select
                Next
                FinishLayout(dgv)
            Finally
                dgv.ResumeLayout()
            End Try
        End Sub

        Public Shared Sub FormatAuditColumns(dgv As DataGridView)
            If dgv Is Nothing OrElse dgv.Columns.Count = 0 Then Return

            Try
                dgv.SuspendLayout()

                For Each col As DataGridViewColumn In dgv.Columns
                    Select Case col.Name
                        Case "AuditID"
                            ConfigureColumn(col, "Audit ID", DataGridViewContentAlignment.MiddleRight)
                        Case "UserID"
                            col.Visible = False
                        Case "UserName"
                            ConfigureColumn(col, "User Account")
                        Case "ActionType"
                            ConfigureColumn(col, "Action Type", DataGridViewContentAlignment.MiddleCenter)
                        Case "ActionDescription"
                            ConfigureColumn(col, "Action Description")
                        Case "Timestamp"
                            ConfigureColumn(col, "Timestamp (UTC)")
                    End Select
                Next
                FinishLayout(dgv)
            Finally
                dgv.ResumeLayout()
            End Try
        End Sub

        Public Shared Sub FormatRoutingColumns(dgv As DataGridView)
            If dgv Is Nothing OrElse dgv.Columns.Count = 0 Then Return

            Try
                dgv.SuspendLayout()

                For Each col As DataGridViewColumn In dgv.Columns
                    Select Case col.Name
                        Case "RoutingID"
                            ConfigureColumn(col, "Log ID", DataGridViewContentAlignment.MiddleRight)
                        Case "DocumentID"
                            col.Visible = False
                        Case "FromOffice"
                            ConfigureColumn(col, "Originating Office")
                        Case "ToOffice"
                            ConfigureColumn(col, "Destination Office")
                        Case "ActionTaken"
                            ConfigureColumn(col, "Action Taken")
                        Case "Remarks"
                            ConfigureColumn(col, "Routing Remarks")
                        Case "RoutedBy"
                            ConfigureColumn(col, "Routed By")
                        Case "Timestamp"
                            ConfigureColumn(col, "Date Transmitted")
                    End Select
                Next
                FinishLayout(dgv)
            Finally
                dgv.ResumeLayout()
            End Try
        End Sub

        Public Shared Sub FormatMovementColumns(dgv As DataGridView)
            If dgv Is Nothing OrElse dgv.Columns.Count = 0 Then Return

            Try
                dgv.SuspendLayout()

                For Each col As DataGridViewColumn In dgv.Columns
                    Select Case col.Name
                        Case "MovementID"
                            ConfigureColumn(col, "Move ID", DataGridViewContentAlignment.MiddleRight)
                        Case "DocumentID"
                            col.Visible = False
                        Case "FromLocation"
                            ConfigureColumn(col, "Prior Storage Landmark")
                        Case "ToLocation"
                            ConfigureColumn(col, "New Storage Landmark")
                        Case "MovedBy"
                            ConfigureColumn(col, "Transferred By")
                        Case "Reason"
                            ConfigureColumn(col, "Transfer Reason / Justification")
                        Case "Timestamp"
                            ConfigureColumn(col, "Date Transferred")
                    End Select
                Next
                FinishLayout(dgv)
            Finally
                dgv.ResumeLayout()
            End Try
        End Sub

        Public Shared Sub FormatPortalQueueColumns(dgv As DataGridView)
            If dgv Is Nothing OrElse dgv.Columns.Count = 0 Then Return

            Try
                dgv.SuspendLayout()

                For Each col As DataGridViewColumn In dgv.Columns
                    Select Case col.Name
                        Case "ControlNumber"
                            ConfigureColumn(col, "External CN", font:=CivicCalmTheme.FontTabular)
                        Case "Category"
                            ConfigureColumn(col, "Category", DataGridViewContentAlignment.MiddleCenter)
                        Case "DocumentTitle"
                            ConfigureColumn(col, "Document Title / Subject")
                        Case "RequesterName"
                            ConfigureColumn(col, "Submitter Name")
                        Case "RequesterEmail"
                            ConfigureColumn(col, "Submitter Email")
                        Case "RequesterPhone"
                            ConfigureColumn(col, "Phone")
                        Case "CreatedAt"
                            ConfigureColumn(col, "Submitted At (UTC)")
                    End Select
                Next
                FinishLayout(dgv)
            Finally
                dgv.ResumeLayout()
            End Try
        End Sub
    End Class
End Namespace
