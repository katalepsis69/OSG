Option Explicit On
Option Strict On

Imports System.Drawing
Imports System.Reflection
Imports System.Windows.Forms

Namespace BTA_OSG
    Public NotInheritable Class DataGridStyler
        Private Sub New()
        End Sub

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

            ' Alternating Row Styling (Soft Off-White Tint #F9FAFB)
            dgv.AlternatingRowsDefaultCellStyle.BackColor = ColorTranslator.FromHtml("#F9FAFB")
            dgv.AlternatingRowsDefaultCellStyle.ForeColor = CivicCalmTheme.ColorInk
            dgv.AlternatingRowsDefaultCellStyle.SelectionBackColor = CivicCalmTheme.ColorPrimarySoft
            dgv.AlternatingRowsDefaultCellStyle.SelectionForeColor = CivicCalmTheme.ColorPrimary
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
                watermarkLabel.Text = "Loading document records..."
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

        Public Shared Sub FormatDocumentColumns(dgv As DataGridView)
            If dgv Is Nothing OrElse dgv.Columns.Count = 0 Then Return

            Try
                dgv.SuspendLayout()
                dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

                For Each col As DataGridViewColumn In dgv.Columns
                    Select Case col.Name
                        Case "DocumentID"
                            col.HeaderText = "ID"
                            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                            col.Width = 55
                            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        Case "DocCode"
                            col.HeaderText = "Document Code"
                            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                            col.Width = 140
                            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                            col.DefaultCellStyle.Font = CivicCalmTheme.FontIdentifier
                        Case "DocType"
                            col.HeaderText = "Classification"
                            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                            col.Width = 140
                        Case "Title"
                            col.HeaderText = "Document Title / Subject"
                            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                        Case "OriginatingOffice"
                            col.HeaderText = "Origin Office"
                            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                            col.Width = 150
                        Case "DestinationOffice"
                            col.HeaderText = "Destination"
                            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                            col.Width = 150
                        Case "CurrentStatus"
                            col.HeaderText = "Status"
                            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                            col.Width = 130
                            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                        Case "AssignedStaff"
                            col.HeaderText = "Assigned Staff"
                            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                            col.Width = 140
                        Case "DateReceived"
                            col.HeaderText = "Date Received"
                            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                            col.Width = 140
                            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        Case "GDriveURL"
                            col.HeaderText = "Soft Copy Link"
                            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                            col.Width = 150
                        Case "CabinetID", "ShelfNo", "BoxCode"
                            col.Visible = False
                    End Select
                Next
            Finally
                dgv.ResumeLayout()
            End Try
        End Sub

        Public Shared Sub FormatDirectiveColumns(dgv As DataGridView)
            If dgv Is Nothing OrElse dgv.Columns.Count = 0 Then Return

            Try
                dgv.SuspendLayout()
                dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

                For Each col As DataGridViewColumn In dgv.Columns
                    Select Case col.Name
                        Case "DirectiveID"
                            col.HeaderText = "ID"
                            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                            col.Width = 50
                            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        Case "DocumentID"
                            col.HeaderText = "Doc ID"
                            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                            col.Width = 60
                            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        Case "SGDirective"
                            col.HeaderText = "Directive Action"
                            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                            col.Width = 180
                            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                        Case "AssignedTo"
                            col.HeaderText = "Assigned Staff"
                            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                            col.Width = 160
                            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                        Case "Notes"
                            col.HeaderText = "Directive Notes / Remarks"
                            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                        Case "LogUser"
                            col.HeaderText = "Issued By"
                            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                            col.Width = 160
                            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                        Case "Timestamp"
                            col.HeaderText = "Date / Time Issued"
                            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                            col.Width = 150
                            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    End Select
                Next
            Finally
                dgv.ResumeLayout()
            End Try
        End Sub

        Public Shared Sub FormatUserColumns(dgv As DataGridView)
            If dgv Is Nothing OrElse dgv.Columns.Count = 0 Then Return

            Try
                dgv.SuspendLayout()
                dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

                For Each col As DataGridViewColumn In dgv.Columns
                    Select Case col.Name
                        Case "UserID"
                            col.HeaderText = "User ID"
                            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                            col.Width = 65
                            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        Case "RFID_UID"
                            col.HeaderText = "RFID Badge UID"
                            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                            col.Width = 140
                            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                            col.DefaultCellStyle.Font = CivicCalmTheme.FontIdentifier
                        Case "FullName"
                            col.HeaderText = "Full Name"
                            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                        Case "Role"
                            col.HeaderText = "Designated Role"
                            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                            col.Width = 200
                            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                        Case "IsActive"
                            col.HeaderText = "Active"
                            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                            col.Width = 70
                            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    End Select
                Next
            Finally
                dgv.ResumeLayout()
            End Try
        End Sub

        Public Shared Sub FormatAuditColumns(dgv As DataGridView)
            If dgv Is Nothing OrElse dgv.Columns.Count = 0 Then Return

            Try
                dgv.SuspendLayout()
                dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

                For Each col As DataGridViewColumn In dgv.Columns
                    Select Case col.Name
                        Case "AuditID"
                            col.HeaderText = "Audit ID"
                            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                            col.Width = 70
                            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        Case "UserName"
                            col.HeaderText = "User Account"
                            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                            col.Width = 180
                            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                        Case "ActionDescription"
                            col.HeaderText = "Action Description"
                            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                        Case "Timestamp"
                            col.HeaderText = "Timestamp (UTC)"
                            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                            col.Width = 160
                            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    End Select
                Next
            Finally
                dgv.ResumeLayout()
            End Try
        End Sub

        Public Shared Sub FormatRoutingColumns(dgv As DataGridView)
            If dgv Is Nothing OrElse dgv.Columns.Count = 0 Then Return

            Try
                dgv.SuspendLayout()
                dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

                For Each col As DataGridViewColumn In dgv.Columns
                    Select Case col.Name
                        Case "RoutingID"
                            col.HeaderText = "Log ID"
                            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                            col.Width = 60
                            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        Case "DocumentID"
                            col.Visible = False
                        Case "FromOffice"
                            col.HeaderText = "Originating Office"
                            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                            col.Width = 150
                            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                        Case "ToOffice"
                            col.HeaderText = "Destination Office"
                            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                            col.Width = 150
                            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                        Case "ActionTaken"
                            col.HeaderText = "Action Taken"
                            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                            col.Width = 140
                            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                        Case "Remarks"
                            col.HeaderText = "Routing Remarks"
                            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                        Case "RoutedBy"
                            col.HeaderText = "Routed By"
                            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                            col.Width = 150
                            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                        Case "Timestamp"
                            col.HeaderText = "Date Transmitted"
                            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                            col.Width = 150
                            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    End Select
                Next
            Finally
                dgv.ResumeLayout()
            End Try
        End Sub

        Public Shared Sub FormatMovementColumns(dgv As DataGridView)
            If dgv Is Nothing OrElse dgv.Columns.Count = 0 Then Return

            Try
                dgv.SuspendLayout()
                dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

                For Each col As DataGridViewColumn In dgv.Columns
                    Select Case col.Name
                        Case "MovementID"
                            col.HeaderText = "Move ID"
                            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                            col.Width = 60
                            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        Case "DocumentID"
                            col.Visible = False
                        Case "FromLocation"
                            col.HeaderText = "Prior Storage Landmark"
                            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                            col.Width = 160
                            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                        Case "ToLocation"
                            col.HeaderText = "New Storage Landmark"
                            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                            col.Width = 160
                            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                        Case "MovedBy"
                            col.HeaderText = "Transferred By"
                            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                            col.Width = 150
                            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                        Case "Reason"
                            col.HeaderText = "Transfer Reason / Justification"
                            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                        Case "Timestamp"
                            col.HeaderText = "Date Transferred"
                            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                            col.Width = 150
                            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    End Select
                Next
            Finally
                dgv.ResumeLayout()
            End Try
        End Sub
    End Class
End Namespace
