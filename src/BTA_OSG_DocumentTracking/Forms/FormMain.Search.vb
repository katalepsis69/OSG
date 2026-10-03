Option Explicit On
Option Strict On

Imports System
Imports System.Data
Imports System.Diagnostics
Imports System.Drawing
Imports System.Windows.Forms

Namespace BTA_OSG
    Partial Public Class FormMain
        Private Sub SetupSearchView()
            viewSearch = New Panel With {.Dock = DockStyle.Fill, .BackColor = CivicCalmTheme.ColorCanvas}

            Dim tblSearch As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .RowCount = 2,
                .BackColor = CivicCalmTheme.ColorCanvas,
                .Padding = New Padding(0)
            }
            tblSearch.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            tblSearch.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            tblSearch.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

            Dim pnlSearchCard As New Panel With {
                .Dock = DockStyle.Fill,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .Margin = New Padding(0, 0, 0, 12),
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(12)
            }
            ApplyCardBorder(pnlSearchCard)

            Dim flwSearch As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .FlowDirection = FlowDirection.LeftToRight,
                .WrapContents = True,
                .BackColor = CivicCalmTheme.ColorSurface
            }

            Dim lblKey As New Label With {
                .Text = "Search Keyword:",
                .UseMnemonic = False,
                .Font = CivicCalmTheme.FontFieldLabel,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .AutoSize = True,
                .Margin = New Padding(0, 8, 8, 0)
            }

            txtSearchKey = New TextBox With {
                .Width = CInt(Dpi(220.0F)),
                .Height = 32,
                .BackColor = CivicCalmTheme.ColorSurface,
                .ForeColor = CivicCalmTheme.ColorInk,
                .BorderStyle = BorderStyle.FixedSingle,
                .Margin = New Padding(0, 4, 8, 0)
            }
            AddHandler txtSearchKey.KeyDown, Sub(s As Object, e As KeyEventArgs)
                                                 If e.KeyCode = Keys.Enter Then
                                                     e.SuppressKeyPress = True
                                                     OnSearch(s, EventArgs.Empty)
                                                 End If
                                             End Sub

            btnSearch = New Button With {
                .Text = "&Search",
                .Size = New Size(85, 32),
                .BackColor = CivicCalmTheme.ColorPrimary,
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(0, 4, 8, 0)
            }
            btnSearch.FlatAppearance.BorderSize = 0

            btnClearSearch = New Button With {
                .Text = "&Clear Filters",
                .Size = New Size(95, 32),
                .BackColor = CivicCalmTheme.ColorWell,
                .ForeColor = CivicCalmTheme.ColorInk,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontBody,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(0, 4, 8, 0)
            }
            btnClearSearch.FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder

            btnViewSearchDetail = New Button With {
                .Text = "&View Details",
                .Size = New Size(110, 32),
                .BackColor = CivicCalmTheme.ColorWell,
                .ForeColor = CivicCalmTheme.ColorInk,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(0, 4, 8, 0)
            }
            btnViewSearchDetail.FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder

            btnOpenPDF = New Button With {
                .Text = "&Launch PDF",
                .Size = New Size(110, 32),
                .BackColor = CivicCalmTheme.ColorWell,
                .ForeColor = CivicCalmTheme.ColorInk,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(0, 4, 0, 0)
            }
            btnOpenPDF.FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder

            AddHandler btnSearch.Click, AddressOf OnSearch
            AddHandler btnClearSearch.Click, Sub()
                                                 txtSearchKey.Text = ""
                                                 OnSearch(Nothing, EventArgs.Empty)
                                             End Sub
            AddHandler btnViewSearchDetail.Click, Sub() OpenSelectedDocumentDetail(dgvSearch)
            AddHandler btnOpenPDF.Click, AddressOf OnOpenPDF

            flwSearch.Controls.AddRange(New Control() {lblKey, txtSearchKey, btnSearch, btnClearSearch, btnViewSearchDetail, btnOpenPDF})
            pnlSearchCard.Controls.Add(flwSearch)

            Dim pnlGridCard As New Panel With {
                .Dock = DockStyle.Fill,
                .Margin = New Padding(0),
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(12)
            }
            ApplyCardBorder(pnlGridCard)

            dgvSearch = New DataGridView With {
                .Dock = DockStyle.Fill,
                .ReadOnly = True,
                .AllowUserToAddRows = False
            }
            DataGridStyler.ApplyCivicStyle(dgvSearch)
            AddHandler dgvSearch.CellDoubleClick, Sub(s, e) If e.RowIndex >= 0 Then OpenSelectedDocumentDetail(dgvSearch)

            lblSearchWatermark = CreateGridWatermark(pnlGridCard)

            pnlGridCard.Controls.Add(dgvSearch)

            tblSearch.Controls.Add(pnlSearchCard, 0, 0)
            tblSearch.Controls.Add(pnlGridCard, 0, 1)

            viewSearch.Controls.Add(tblSearch)
        End Sub

        Public Shared Function EscapeRowFilter(raw As String) As String
            If String.IsNullOrEmpty(raw) Then Return ""
            Dim sb As New System.Text.StringBuilder()
            For Each ch As Char In raw
                Select Case ch
                    Case "'"c
                        sb.Append("''")
                    Case "["c
                        sb.Append("[[]")
                    Case "]"c
                        sb.Append("[]]")
                    Case "*"c
                        sb.Append("[*]")
                    Case "%"c
                        sb.Append("[%]")
                    Case Else
                        sb.Append(ch)
                End Select
            Next
            Return sb.ToString()
        End Function

        Private Sub OnSearch(sender As Object, e As EventArgs)
            Dim visibleView = EmbeddedDB.GetVisibleDocuments(CurrentUser)
            Dim rawQuery = txtSearchKey.Text.Trim()
            If String.IsNullOrEmpty(rawQuery) Then
                dgvSearch.DataSource = visibleView
                lblStatusMessage.Text = "Displaying all visible documents."
            Else
                Try
                    Dim safeQ = EscapeRowFilter(rawQuery)
                    ' The DataView from GetVisibleDocuments is created per call, so setting its
                    ' RowFilter here cannot leak the filter into the dashboard or registry grids.
                    visibleView.RowFilter = String.Format("DocCode LIKE '%{0}%' OR Title LIKE '%{0}%' OR DocType LIKE '%{0}%' OR CabinetID LIKE '%{0}%' OR OriginatingOffice LIKE '%{0}%' OR DestinationOffice LIKE '%{0}%' OR AssignedStaff LIKE '%{0}%'", safeQ)
                    dgvSearch.DataSource = visibleView
                    lblStatusMessage.Text = String.Format("Search complete for '{0}'.", rawQuery)
                Catch ex As Exception
                    dgvSearch.DataSource = visibleView.Table.Clone()
                    lblStatusMessage.Text = "Invalid search filter syntax entered."
                End Try
            End If
            DataGridStyler.FormatDocumentColumns(dgvSearch)

            If dgvSearch.Rows.Count = 0 Then
                DataGridStyler.SetEmptyState(dgvSearch, lblSearchWatermark, "No documents match the current filter criteria. Press Alt+C to clear filters.")
            Else
                DataGridStyler.SetPopulatedState(dgvSearch, lblSearchWatermark)
            End If
            lblStatusCount.Text = dgvSearch.Rows.Count.ToString() & " Records"
        End Sub

        Private Sub OnOpenPDF(sender As Object, e As EventArgs)
            If dgvSearch.CurrentRow Is Nothing Then
                lblStatusMessage.Text = "Please select a document row from the grid first."
                Return
            End If

            ' Deny-by-default: only an explicit CanSoftCopy=True grants soft-copy launch.
            Dim canSoftCopy As Boolean = False
            If CurrentUser IsNot Nothing AndAlso CurrentUser.Table.Columns.Contains("CanSoftCopy") AndAlso Not IsDBNull(CurrentUser("CanSoftCopy")) Then
                canSoftCopy = Convert.ToBoolean(CurrentUser("CanSoftCopy"))
            End If
            If Not canSoftCopy Then
                lblStatusMessage.Text = "Access Denied: Current user does not have permission to view or launch soft-copy attachments."
                Return
            End If

            Dim url As String = Convert.ToString(dgvSearch.CurrentRow.Cells("GDriveURL").Value)
            If Not String.IsNullOrEmpty(url) Then
                Dim errUrl As String = ""
                If Not EmbeddedDB.ValidateGDriveURL(url, errUrl) Then
                    lblStatusMessage.Text = "Security Error: " & errUrl
                    Return
                End If
                Try
                    Process.Start(New ProcessStartInfo With {.FileName = url, .UseShellExecute = True})
                    If CurrentUser IsNot Nothing Then
                        EmbeddedDB.LogAudit(CurrentUser("FullName").ToString(), "Launched Document Soft Copy: " & url)
                    End If
                    lblStatusMessage.Text = "Launched document in default viewer."
                Catch ex As Exception
                    lblStatusMessage.Text = "Unable to launch document viewer: " & ex.Message
                End Try
            Else
                lblStatusMessage.Text = "No digital soft copy attached to this document."
            End If
        End Sub
    End Class
End Namespace
