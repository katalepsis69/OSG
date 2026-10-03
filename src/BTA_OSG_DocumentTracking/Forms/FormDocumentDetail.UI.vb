Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Drawing
Imports System.Windows.Forms

Namespace BTA_OSG
    Partial Public Class FormDocumentDetail
        Private Sub SetupHeaderPanel()
            pnlHeader = New Panel With {
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .MinimumSize = New Size(0, 135),
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(20, 14, 20, 14)
            }

            Dim pnlBorderBottom As New Panel With {
                .Dock = DockStyle.Bottom,
                .Height = 1,
                .BackColor = CivicCalmTheme.ColorBorder
            }
            pnlHeader.Controls.Add(pnlBorderBottom)

            Dim tblHeader As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 2,
                .RowCount = 1,
                .BackColor = CivicCalmTheme.ColorSurface
            }
            tblHeader.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 55.0F))
            tblHeader.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 45.0F))

            Dim pnlTitles As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.TopDown,
                .WrapContents = False,
                .BackColor = CivicCalmTheme.ColorSurface
            }

            lblCode = New Label With {
                .Text = String.Format("DOC CODE: {0}   |   TYPE: {1}   |   FLOW: {2}", DocRow("DocCode"), DocRow("DocType").ToString().ToUpperInvariant(), DocRow("FlowDirection")),
                .Font = CivicCalmTheme.FontFormTitle,
                .ForeColor = CivicCalmTheme.ColorPrimary,
                .AutoSize = True,
                .Margin = New Padding(0, 0, 0, 4)
            }

            lblTitle = New Label With {
                .Text = DocRow("Title").ToString(),
                .Font = CivicCalmTheme.FontSectionHeader,
                .ForeColor = CivicCalmTheme.ColorInk,
                .AutoSize = True,
                .Margin = New Padding(0, 0, 0, 4)
            }

            lblStatusBadge = New Label With {
                .Text = String.Format("Status: {0}   |   Section: {1}   |   Staff: {2}", DocumentStatus.DisplayName(DocRow("CurrentStatus").ToString()), DocRow("AssignedSection"), DocRow("AssignedStaff")),
                .Font = CivicCalmTheme.FontFieldLabel,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .AutoSize = True
            }
            pnlTitles.Controls.AddRange(New Control() {lblCode, lblTitle, lblStatusBadge})

            Dim pnlActions = CreateActionsPanel()

            tblHeader.Controls.Add(pnlTitles, 0, 0)
            tblHeader.Controls.Add(pnlActions, 1, 0)
            pnlHeader.Controls.Add(tblHeader)
        End Sub

        Private Function CreateActionsPanel() As FlowLayoutPanel
            Dim pnlActions As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.RightToLeft,
                .WrapContents = True,
                .BackColor = CivicCalmTheme.ColorSurface
            }

            btnClose = New Button With {
                .Text = "&Close",
                .Size = New Size(80, 34),
                .BackColor = CivicCalmTheme.ColorWell,
                .ForeColor = CivicCalmTheme.ColorInk,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(4, 2, 0, 2)
            }
            btnClose.FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder
            AddHandler btnClose.Click, Sub() Me.Close()

            btnLaunchPdf = New Button With {
                .Text = "&Open Link",
                .Size = New Size(95, 34),
                .BackColor = CivicCalmTheme.ColorWell,
                .ForeColor = CivicCalmTheme.ColorInk,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(4, 2, 0, 2)
            }
            btnLaunchPdf.FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder
            AddHandler btnLaunchPdf.Click, AddressOf OnLaunchPDF

            btnMove = New Button With {
                .Text = "&Transfer Storage",
                .Size = New Size(130, 34),
                .BackColor = CivicCalmTheme.ColorWell,
                .ForeColor = CivicCalmTheme.ColorInk,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(4, 2, 0, 2)
            }
            btnMove.FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder
            AddHandler btnMove.Click, AddressOf OnMoveStorage

            btnRoute = New Button With {
                .Text = "&Route",
                .Size = New Size(80, 34),
                .BackColor = CivicCalmTheme.ColorWell,
                .ForeColor = CivicCalmTheme.ColorInk,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(4, 2, 0, 2)
            }
            btnRoute.FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder
            AddHandler btnRoute.Click, AddressOf OnRouteDocument

            btnRelease = New Button With {
                .Text = "Re&lease Doc",
                .Size = New Size(115, 34),
                .BackColor = CivicCalmTheme.ColorAccentSG,
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(4, 2, 0, 2)
            }
            btnRelease.FlatAppearance.BorderSize = 0
            AddHandler btnRelease.Click, AddressOf OnRelease

            btnApprove = New Button With {
                .Text = "&Approve Doc",
                .Size = New Size(115, 34),
                .BackColor = CivicCalmTheme.ColorPrimary,
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(4, 2, 0, 2)
            }
            btnApprove.FlatAppearance.BorderSize = 0
            AddHandler btnApprove.Click, AddressOf OnApprove

            btnResubmit = New Button With {
                .Text = "Re&submit",
                .Size = New Size(100, 34),
                .BackColor = CivicCalmTheme.ColorInfo,
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(4, 2, 0, 2)
            }
            btnResubmit.FlatAppearance.BorderSize = 0
            AddHandler btnResubmit.Click, AddressOf OnResubmit

            btnRequestRevision = New Button With {
                .Text = "Request Re&vision",
                .Size = New Size(135, 34),
                .BackColor = CivicCalmTheme.ColorDanger,
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(4, 2, 0, 2)
            }
            btnRequestRevision.FlatAppearance.BorderSize = 0
            AddHandler btnRequestRevision.Click, AddressOf OnRequestRevision

            btnRoutingSlip = New Button With {
                .Text = "&Print Routing Slip",
                .Size = New Size(145, 34),
                .BackColor = CivicCalmTheme.ColorWell,
                .ForeColor = CivicCalmTheme.ColorInk,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(4, 2, 0, 2)
            }
            btnRoutingSlip.FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder
            AddHandler btnRoutingSlip.Click, AddressOf OnPrintRoutingSlip

            Me.CancelButton = btnClose

            pnlActions.Controls.AddRange(New Control() {btnClose, btnRoutingSlip, btnLaunchPdf, btnMove, btnRoute, btnRelease, btnApprove, btnResubmit, btnRequestRevision})
            Return pnlActions
        End Function

        Private Sub SetupPunchlistBanner()
            pnlPunchlistBanner = New Panel With {
                .Dock = DockStyle.Top,
                .Height = 70,
                .BackColor = CivicCalmTheme.ColorStatusRevisionBg,
                .Padding = New Padding(16, 8, 16, 8),
                .Visible = False
            }

            Dim pnlPunchBorder As New Panel With {
                .Dock = DockStyle.Bottom,
                .Height = 1,
                .BackColor = CivicCalmTheme.ColorPunchlistBorder
            }
            pnlPunchlistBanner.Controls.Add(pnlPunchBorder)

            Dim lblPunchTitle As New Label With {
                .Text = "ACTION REQUIRED: SECRETARY-GENERAL REVISION PUNCHLIST",
                .Font = CivicCalmTheme.FontFieldLabel,
                .ForeColor = CivicCalmTheme.ColorStatusRevisionFg,
                .Dock = DockStyle.Top,
                .Height = 20
            }

            lblPunchlistContent = New Label With {
                .Font = CivicCalmTheme.FontBody,
                .ForeColor = CivicCalmTheme.ColorPunchlistContentFg,
                .Dock = DockStyle.Fill,
                .AutoEllipsis = False
            }
            pnlPunchlistBanner.Controls.AddRange(New Control() {lblPunchlistContent, lblPunchTitle})
        End Sub

        Private Sub SetupTabControl()
            tabDetail = New TabControl With {
                .Dock = DockStyle.Fill,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Padding = New Point(12, 6)
            }

            tabOverview = New TabPage(" Document Overview & Specifications ")
            tabDirectives = New TabPage(" SG Directives Timeline ")
            tabRouting = New TabPage(" Step-by-Step Route & Transmittal Logs ")
            tabMovements = New TabPage(" Storage & Custody Movement History ")
            tabPreview = New TabPage(" Digital Soft-Copy Preview ")

            tabOverview.BackColor = CivicCalmTheme.ColorSurface
            tabDirectives.BackColor = CivicCalmTheme.ColorSurface
            tabRouting.BackColor = CivicCalmTheme.ColorSurface
            tabMovements.BackColor = CivicCalmTheme.ColorSurface
            tabPreview.BackColor = CivicCalmTheme.ColorSurface

            SetupOverviewTab()
            SetupPreviewTab()

            dgvDirectives = CreateDetailGrid()
            dgvMovements = CreateDetailGrid()

            tabDirectives.Controls.Add(dgvDirectives)
            lblDirectivesWatermark = CreateWatermarkLabel()
            tabDirectives.Controls.Add(lblDirectivesWatermark)
            SetupRoutingTab()
            tabMovements.Controls.Add(dgvMovements)
            lblMovementsWatermark = CreateWatermarkLabel()
            tabMovements.Controls.Add(lblMovementsWatermark)

            tabDetail.TabPages.AddRange(New TabPage() {tabOverview, tabDirectives, tabRouting, tabMovements, tabPreview})
        End Sub

        Private Sub SetupWorkflowRibbon()
            pnlWorkflowRibbon = New Panel With {
                .Dock = DockStyle.Top,
                .Height = 44,
                .BackColor = Color.FromArgb(240, 244, 248),
                .Padding = New Padding(20, 6, 20, 6)
            }

            Dim pnlBorderBottom As New Panel With {
                .Dock = DockStyle.Bottom,
                .Height = 1,
                .BackColor = CivicCalmTheme.ColorBorder
            }
            pnlWorkflowRibbon.Controls.Add(pnlBorderBottom)

            Dim tblRibbon As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 3,
                .RowCount = 1,
                .BackColor = Color.FromArgb(240, 244, 248)
            }
            tblRibbon.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 45.0F))
            tblRibbon.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 40.0F))
            tblRibbon.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 15.0F))

            lblStationBadge = New Label With {
                .Text = "CURRENT STATION: Loading",
                .Font = CivicCalmTheme.FontFieldLabel,
                .ForeColor = CivicCalmTheme.ColorPrimary,
                .Dock = DockStyle.Fill,
                .TextAlign = ContentAlignment.MiddleLeft,
                .UseMnemonic = False
            }

            lblNextStationBadge = New Label With {
                .Text = "Next Step: Loading",
                .Font = CivicCalmTheme.FontBody,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .Dock = DockStyle.Fill,
                .TextAlign = ContentAlignment.MiddleLeft,
                .UseMnemonic = False
            }

            lblProgressBadge = New Label With {
                .Text = "0% Complete",
                .Font = CivicCalmTheme.FontFieldLabel,
                .ForeColor = CivicCalmTheme.ColorPrimary,
                .Dock = DockStyle.Fill,
                .TextAlign = ContentAlignment.MiddleRight,
                .UseMnemonic = False
            }

            tblRibbon.Controls.Add(lblStationBadge, 0, 0)
            tblRibbon.Controls.Add(lblNextStationBadge, 1, 0)
            tblRibbon.Controls.Add(lblProgressBadge, 2, 0)
            pnlWorkflowRibbon.Controls.Add(tblRibbon)
        End Sub

        Private Sub SetupRoutingTab()
            tabRouting.Controls.Clear()

            Dim splitRouting As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .RowCount = 4,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(12)
            }
            splitRouting.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            splitRouting.RowStyles.Add(New RowStyle(SizeType.Absolute, 28.0F))
            splitRouting.RowStyles.Add(New RowStyle(SizeType.Percent, 55.0F))
            splitRouting.RowStyles.Add(New RowStyle(SizeType.Absolute, 32.0F))
            splitRouting.RowStyles.Add(New RowStyle(SizeType.Percent, 45.0F))

            Dim lblStepsHdr As New Label With {
                .Text = "Official Step-by-Step Custody & Routing Roadmap",
                .Font = CivicCalmTheme.FontFieldLabel,
                .ForeColor = CivicCalmTheme.ColorPrimary,
                .Dock = DockStyle.Fill,
                .TextAlign = ContentAlignment.MiddleLeft,
                .UseMnemonic = False
            }

            dgvSteps = CreateDetailGrid()

            Dim lblLogsHdr As New Label With {
                .Text = "Chronological Inter-Office Transit Logs",
                .Font = CivicCalmTheme.FontFieldLabel,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .Dock = DockStyle.Fill,
                .TextAlign = ContentAlignment.BottomLeft,
                .UseMnemonic = False
            }

            dgvRouting = CreateDetailGrid()

            splitRouting.Controls.Add(lblStepsHdr, 0, 0)
            splitRouting.Controls.Add(dgvSteps, 0, 1)
            lblStepsWatermark = CreateWatermarkLabel()
            splitRouting.Controls.Add(lblStepsWatermark, 0, 1)
            splitRouting.Controls.Add(lblLogsHdr, 0, 2)
            splitRouting.Controls.Add(dgvRouting, 0, 3)
            lblRoutingWatermark = CreateWatermarkLabel()
            splitRouting.Controls.Add(lblRoutingWatermark, 0, 3)

            tabRouting.Controls.Add(splitRouting)
        End Sub

        Private Sub SetupOverviewTab()
            tabOverview.Controls.Clear()

            Dim pnlScroll As New Panel With {
                .Dock = DockStyle.Fill,
                .AutoScroll = True,
                .Padding = New Padding(20),
                .BackColor = CivicCalmTheme.ColorSurface
            }

            pnlOverviewTable = New TableLayoutPanel With {
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .ColumnCount = 2,
                .CellBorderStyle = TableLayoutPanelCellBorderStyle.Single,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(10)
            }

            pnlOverviewTable.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 28.0F))
            pnlOverviewTable.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 72.0F))

            AddOverviewRow("Document System ID:", "#" & DocRow("DocumentID").ToString())
            AddOverviewRow("Document Code:", DocRow("DocCode").ToString())
            If DocRow.Table.Columns.Contains("ExternalControlNumber") AndAlso Not String.IsNullOrWhiteSpace(DocRow("ExternalControlNumber").ToString()) Then
                AddOverviewRow("External Control Number:", DocRow("ExternalControlNumber").ToString())
            End If
            AddOverviewRow("Classification Category:", DocRow("DocType").ToString())
            AddOverviewRow("Flow Direction:", DocRow("FlowDirection").ToString())
            AddOverviewRow("Assigned Section Desk:", DocRow("AssignedSection").ToString())
            AddOverviewRow("Official Document Title:", DocRow("Title").ToString())
            AddOverviewRow("Originating Office:", DocRow("OriginatingOffice").ToString())
            AddOverviewRow("Destination Office:", DocRow("DestinationOffice").ToString())
            AddOverviewRow("Date Received / Registered:", FormatStoredDate(DocRow("DateReceived").ToString()))
            AddOverviewRow("Target Deadline:", FormatStoredDate(DocRow("TargetDeadlineUTC").ToString()))
            AddOverviewRow("Current Processing Status:", DocumentStatus.DisplayName(DocRow("CurrentStatus").ToString()))
            AddOverviewRow("Assigned OSG Staff Member:", DocRow("AssignedStaff").ToString())
            AddOverviewRow("Last Action Taken:", DocRow("LastActionTaken").ToString())
            AddOverviewRow("Revision Punchlist Notes:", DocRow("RevisionPunchlist").ToString())
            AddOverviewRow("Cabinet Landmark ID:", DocRow("CabinetID").ToString())
            AddOverviewRow("Shelf Landmark No:", DocRow("ShelfNo").ToString())
            AddOverviewRow("Box Landmark Code:", DocRow("BoxCode").ToString())
            Dim storageParts As New List(Of String)()
            If Not String.IsNullOrWhiteSpace(DocRow("CabinetID").ToString()) Then storageParts.Add("Cabinet " & DocRow("CabinetID").ToString())
            If Not String.IsNullOrWhiteSpace(DocRow("ShelfNo").ToString()) Then storageParts.Add("Shelf " & DocRow("ShelfNo").ToString())
            If Not String.IsNullOrWhiteSpace(DocRow("BoxCode").ToString()) Then storageParts.Add("Box " & DocRow("BoxCode").ToString())
            AddOverviewRow("Full Storage String:", String.Join(" / ", storageParts))
            AddOverviewRow("Soft Copy Reference URL:", WrapSoftCopyLink(DocRow("GDriveURL").ToString()))

            pnlScroll.Controls.Add(pnlOverviewTable)
            tabOverview.Controls.Add(pnlScroll)
        End Sub

        Private Sub SetupPreviewTab()
            tabPreview.Controls.Clear()

            Dim pnlPreviewContainer As New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = CivicCalmTheme.ColorCanvas,
                .Padding = New Padding(12)
            }

            Dim pnlTopBar As New FlowLayoutPanel With {
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .BackColor = CivicCalmTheme.ColorSurface,
                .FlowDirection = FlowDirection.LeftToRight,
                .WrapContents = True,
                .Padding = New Padding(8, 6, 8, 6)
            }

            ' Deny-by-default so the external-browser escape hatch obeys the same gate as
            ' the preview tab and the launch button.
            Dim canSoftCopyUser As Boolean = False
            If MainFrm.CurrentUser IsNot Nothing AndAlso MainFrm.CurrentUser.Table.Columns.Contains("CanSoftCopy") AndAlso Not IsDBNull(MainFrm.CurrentUser("CanSoftCopy")) Then
                canSoftCopyUser = Convert.ToBoolean(MainFrm.CurrentUser("CanSoftCopy"))
            End If

            Dim btnOpenExt As New Button With {
                .Text = "&Open in External Browser",
                .Size = New Size(180, 28),
                .Font = CivicCalmTheme.FontMicrocopy,
                .BackColor = CivicCalmTheme.ColorWell,
                .ForeColor = CivicCalmTheme.ColorInk,
                .FlatStyle = FlatStyle.Flat,
                .Enabled = canSoftCopyUser
            }
            btnOpenExt.FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder
            AddHandler btnOpenExt.Click, AddressOf OnLaunchPDF

            Dim lblUrlInfo As New Label With {
                .Text = "Source: " & WrapSoftCopyLink(DocRow("GDriveURL").ToString()),
                .Font = CivicCalmTheme.FontMicrocopy,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .AutoSize = True,
                .Margin = New Padding(12, 6, 0, 0)
            }
            pnlTopBar.Controls.AddRange(New Control() {btnOpenExt, lblUrlInfo})

            wvPreview = New Microsoft.Web.WebView2.WinForms.WebView2 With {
                .Dock = DockStyle.Fill,
                .DefaultBackgroundColor = Color.White
            }

            lblPreviewPlaceholder = New Label With {
                .Text = "No digital soft copy attached or URL unavailable for this document.",
                .Font = CivicCalmTheme.FontSectionHeader,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .TextAlign = ContentAlignment.MiddleCenter,
                .Dock = DockStyle.Fill,
                .Visible = False
            }

            pnlPreviewContainer.Controls.Add(wvPreview)
            pnlPreviewContainer.Controls.Add(lblPreviewPlaceholder)
            pnlPreviewContainer.Controls.Add(pnlTopBar)

            tabPreview.Controls.Add(pnlPreviewContainer)
        End Sub

        Private Sub AddOverviewRow(label As String, value As String)
            Dim lblField As New Label With {
                .Text = label,
                .Font = CivicCalmTheme.FontFieldLabel,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .Dock = DockStyle.Fill,
                .TextAlign = ContentAlignment.MiddleLeft,
                .Padding = New Padding(8),
                .AutoSize = True
            }

            Dim lblVal As New Label With {
                .Text = If(String.IsNullOrWhiteSpace(value), "(None Specified)", value),
                .Font = CivicCalmTheme.FontBody,
                .ForeColor = CivicCalmTheme.ColorInk,
                .Dock = DockStyle.Fill,
                .TextAlign = ContentAlignment.MiddleLeft,
                .Padding = New Padding(8),
                .AutoSize = True
            }

            Dim rowIndex = pnlOverviewTable.RowCount
            pnlOverviewTable.RowCount += 1
            pnlOverviewTable.Controls.Add(lblField, 0, rowIndex)
            pnlOverviewTable.Controls.Add(lblVal, 1, rowIndex)
        End Sub

        ' Dates are stored as free-form strings from the registration picker; render them in
        ' the system's tabular stamp format when they parse, verbatim when they do not.
        Private Shared Function FormatStoredDate(raw As String) As String
            Dim parsed As DateTime
            If DateTime.TryParse(raw, parsed) Then Return parsed.ToString("yyyy-MM-dd HH:mm")
            Return raw
        End Function

        ' A Label cannot break an unbroken URL token, so insert soft breaks at path separators
        ' instead of letting a long link clip mid-string.
        Public Shared Function WrapSoftCopyLink(url As String) As String
            If String.IsNullOrWhiteSpace(url) Then Return ""

            Const segmentLimit As Integer = 48
            Dim sb As New System.Text.StringBuilder(url.Length + 8)
            Dim segmentLength As Integer = 0
            For Each ch As Char In url
                sb.Append(ch)
                segmentLength += 1
                If ch = "/"c OrElse ch = "?"c OrElse ch = "&"c OrElse ch = "="c OrElse segmentLength >= segmentLimit Then
                    sb.Append(Environment.NewLine)
                    segmentLength = 0
                End If
            Next
            Return sb.ToString().TrimEnd()
        End Function

        Private Function CreateDetailGrid() As DataGridView
            Dim dgv As New DataGridView With {
                .Dock = DockStyle.Fill,
                .ReadOnly = True,
                .AllowUserToAddRows = False
            }
            DataGridStyler.ApplyCivicStyle(dgv)
            Return dgv
        End Function

        Private Function CreateWatermarkLabel() As Label
            Return New Label With {
                .Dock = DockStyle.Fill,
                .TextAlign = ContentAlignment.MiddleCenter,
                .Visible = False,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .Font = CivicCalmTheme.FontBody,
                .BackColor = CivicCalmTheme.ColorSurface
            }
        End Function

        Private Sub ShowGridState(dgv As DataGridView, watermark As Label, emptyMessage As String)
            If dgv.Rows.Count = 0 Then
                DataGridStyler.SetEmptyState(dgv, watermark, emptyMessage)
            Else
                DataGridStyler.SetPopulatedState(dgv, watermark)
            End If
        End Sub

        Private Sub PopulateStepsGrid(steps As List(Of RoutingStep))
            Dim dt As New DataTable()
            dt.Columns.Add("StepNo", GetType(Integer))
            dt.Columns.Add("Stage", GetType(String))
            dt.Columns.Add("ResponsibleOffice", GetType(String))
            dt.Columns.Add("ActionRequired", GetType(String))
            dt.Columns.Add("Status", GetType(String))
            dt.Columns.Add("CompletedAt", GetType(String))
            dt.Columns.Add("Remarks", GetType(String))

            For Each s In steps
                Dim statusText = s.StepStatus
                If s.StepStatus = "CURRENT" Then
                    statusText = ">>> CURRENT STATION <<<"
                End If
                Dim completedStr = If(s.CompletedAtUTC.HasValue, s.CompletedAtUTC.Value.ToString("yyyy-MM-dd HH:mm"), "")
                dt.Rows.Add(s.StepNumber, s.StageName, s.ResponsibleOffice, s.ActionRequired, statusText, completedStr, s.Remarks)
            Next

            dgvSteps.DataSource = dt
            ShowGridState(dgvSteps, lblStepsWatermark, "Pipeline steps appear here once the document is registered.")

            If dgvSteps.Columns.Contains("StepNo") Then
                dgvSteps.Columns("StepNo").HeaderText = "Step"
                dgvSteps.Columns("StepNo").MinimumWidth = 60
                dgvSteps.Columns("StepNo").FillWeight = 40
                dgvSteps.Columns("StepNo").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            End If
            If dgvSteps.Columns.Contains("Stage") Then
                dgvSteps.Columns("Stage").HeaderText = "Stage / Station"
                dgvSteps.Columns("Stage").MinimumWidth = 190
                dgvSteps.Columns("Stage").FillWeight = 140
                dgvSteps.Columns("Stage").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            End If
            If dgvSteps.Columns.Contains("ResponsibleOffice") Then
                dgvSteps.Columns("ResponsibleOffice").HeaderText = "Designated Desk"
                dgvSteps.Columns("ResponsibleOffice").MinimumWidth = 160
                dgvSteps.Columns("ResponsibleOffice").FillWeight = 120
                dgvSteps.Columns("ResponsibleOffice").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            End If
            If dgvSteps.Columns.Contains("ActionRequired") Then
                dgvSteps.Columns("ActionRequired").HeaderText = "Action Required"
                dgvSteps.Columns("ActionRequired").MinimumWidth = 240
                dgvSteps.Columns("ActionRequired").FillWeight = 260
                dgvSteps.Columns("ActionRequired").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                dgvSteps.Columns("ActionRequired").DefaultCellStyle.WrapMode = DataGridViewTriState.True
            End If
            If dgvSteps.Columns.Contains("Status") Then
                dgvSteps.Columns("Status").HeaderText = "Progress Status"
                dgvSteps.Columns("Status").MinimumWidth = 160
                dgvSteps.Columns("Status").FillWeight = 130
                dgvSteps.Columns("Status").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            End If
            If dgvSteps.Columns.Contains("CompletedAt") Then
                dgvSteps.Columns("CompletedAt").HeaderText = "Completed Date"
                dgvSteps.Columns("CompletedAt").MinimumWidth = 130
                dgvSteps.Columns("CompletedAt").FillWeight = 100
                dgvSteps.Columns("CompletedAt").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            End If
            If dgvSteps.Columns.Contains("Remarks") Then
                dgvSteps.Columns("Remarks").HeaderText = "Notes / Custody Remarks"
                dgvSteps.Columns("Remarks").MinimumWidth = 160
                dgvSteps.Columns("Remarks").FillWeight = 150
                dgvSteps.Columns("Remarks").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                dgvSteps.Columns("Remarks").DefaultCellStyle.WrapMode = DataGridViewTriState.True
            End If

            For Each row As DataGridViewRow In dgvSteps.Rows
                If row.Cells("Status").Value IsNot Nothing Then
                    Dim st = row.Cells("Status").Value.ToString()
                    If st.Contains("CURRENT") Then
                        row.DefaultCellStyle.BackColor = Color.FromArgb(232, 244, 253)
                        row.DefaultCellStyle.Font = CivicCalmTheme.FontFieldLabel
                        row.DefaultCellStyle.ForeColor = Color.FromArgb(13, 71, 161)
                    ElseIf st = "COMPLETED" Then
                        row.Cells("Status").Style.ForeColor = Color.FromArgb(46, 125, 50)
                    Else
                        row.Cells("Status").Style.ForeColor = Color.FromArgb(100, 116, 139)
                    End If
                End If
            Next
            dgvSteps.ClearSelection()
            dgvSteps.AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCells)
        End Sub
    End Class
End Namespace
