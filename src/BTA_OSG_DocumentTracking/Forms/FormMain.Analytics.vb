Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Globalization
Imports System.Windows.Forms

Namespace BTA_OSG
    Partial Public Class FormMain
        ' Analytics View Controls
        Private cmbAnalyticsTimeframe As ComboBox
        Private cmbAnalyticsSection As ComboBox
        Private btnAnalyticsRefresh As Button
        Private lblAnalyticsPeriodNotice As Label

        Private lblAnalyticsStatTotal As Label
        Private lblAnalyticsStatApproved As Label
        Private lblAnalyticsStatPending As Label
        Private lblAnalyticsStatVelocity As Label

        Private dgvMostRequested As DataGridView
        Private lblMostRequestedWatermark As Label
        Private pnlMostReqChart As Panel

        Private dgvSectionBreakdown As DataGridView
        Private lblSectionWatermark As Label
        Private pnlSectionChart As Panel

        Private pnlGadDemographics As Panel
        Private lblGadTotalSubmissions As Label
        Private lblGadFemaleCount As Label
        Private lblGadMaleCount As Label
        Private lblGadUndisclosedCount As Label
        Private pnlGadRatioBar As Panel

        Private _gadFemaleRatio As Single = 0.0F
        Private _gadMaleRatio As Single = 0.0F
        Private _gadOtherRatio As Single = 0.0F

        Private _categoryChartItems As New List(Of CategoryChartItem)()
        Private _sectionChartItems As New List(Of SectionDeskStats)()

        Private Sub SetupAnalyticsView()
            viewAnalytics = New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = CivicCalmTheme.ColorCanvas,
                .AutoScroll = True,
                .AutoScrollMinSize = New Size(CInt(Dpi(960.0F)), CInt(Dpi(760.0F)))
            }

            Dim tblMain As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .RowCount = 3,
                .BackColor = CivicCalmTheme.ColorCanvas,
                .Padding = New Padding(0),
                .MinimumSize = New Size(CInt(Dpi(960.0F)), CInt(Dpi(760.0F)))
            }
            tblMain.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            tblMain.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            tblMain.RowStyles.Add(New RowStyle(SizeType.Absolute, 96.0F))
            tblMain.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

            ' 1. Filter & Controls Card
            Dim pnlFilterCard As New Panel With {
                .Dock = DockStyle.Fill,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .Margin = New Padding(0, 0, 0, 10),
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(12)
            }
            ApplyCardBorder(pnlFilterCard)

            Dim flwFilter As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .FlowDirection = FlowDirection.LeftToRight,
                .WrapContents = True,
                .BackColor = CivicCalmTheme.ColorSurface
            }

            Dim lblTf As New Label With {
                .Text = "Reporting Period:",
                .Font = CivicCalmTheme.FontFieldLabel,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .AutoSize = True,
                .Margin = New Padding(0, 7, 4, 0)
            }
            cmbAnalyticsTimeframe = New ComboBox With {
                .DropDownStyle = ComboBoxStyle.DropDownList,
                .Width = CInt(Dpi(160.0F)),
                .Height = 28,
                .Font = CivicCalmTheme.FontBody,
                .BackColor = CivicCalmTheme.ColorSurface,
                .ForeColor = CivicCalmTheme.ColorInk,
                .Margin = New Padding(0, 3, 16, 0)
            }
            cmbAnalyticsTimeframe.Items.AddRange(New Object() {"Today (Day)", "This Week (7 Days)", "This Month (30 Days)", "This Year (Year)", "All Time"})
            cmbAnalyticsTimeframe.SelectedIndex = 2 ' Default: This Month
            AddHandler cmbAnalyticsTimeframe.SelectedIndexChanged, Sub() RefreshAnalyticsView()

            Dim lblSec As New Label With {
                .Text = "Section Desk:",
                .Font = CivicCalmTheme.FontFieldLabel,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .AutoSize = True,
                .Margin = New Padding(0, 7, 4, 0)
            }
            cmbAnalyticsSection = New ComboBox With {
                .DropDownStyle = ComboBoxStyle.DropDownList,
                .Width = CInt(Dpi(180.0F)),
                .Height = 28,
                .Font = CivicCalmTheme.FontBody,
                .BackColor = CivicCalmTheme.ColorSurface,
                .ForeColor = CivicCalmTheme.ColorInk,
                .Margin = New Padding(0, 3, 16, 0)
            }
            cmbAnalyticsSection.Items.AddRange(New Object() {"All Sections", "Records Section", "Secretariat", "Legislative Section", "Finance Section", "Travel Section"})
            cmbAnalyticsSection.SelectedIndex = 0
            AddHandler cmbAnalyticsSection.SelectedIndexChanged, Sub() RefreshAnalyticsView()

            btnAnalyticsRefresh = New Button With {
                .Text = "&Refresh Analytics",
                .Size = New Size(140, 32),
                .BackColor = CivicCalmTheme.ColorPrimary,
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(0, 1, 16, 0)
            }
            btnAnalyticsRefresh.FlatAppearance.BorderSize = 0
            AddHandler btnAnalyticsRefresh.Click, Sub() RefreshAnalyticsView()

            lblAnalyticsPeriodNotice = New Label With {
                .Text = "Analytics Period: Current Month",
                .Font = CivicCalmTheme.FontMicrocopy,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .AutoSize = True,
                .Margin = New Padding(0, 8, 0, 0)
            }

            flwFilter.Controls.AddRange(New Control() {lblTf, cmbAnalyticsTimeframe, lblSec, cmbAnalyticsSection, btnAnalyticsRefresh, lblAnalyticsPeriodNotice})
            pnlFilterCard.Controls.Add(flwFilter)

            ' 2. KPI Cards Row
            Dim pnlCards As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .Margin = New Padding(0, 0, 0, 10),
                .ColumnCount = 4,
                .RowCount = 1,
                .BackColor = CivicCalmTheme.ColorCanvas
            }
            pnlCards.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 25.0F))
            pnlCards.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 25.0F))
            pnlCards.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 25.0F))
            pnlCards.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 25.0F))
            pnlCards.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

            lblAnalyticsStatTotal = CreateStatCard(pnlCards, 0, "TOTAL INGESTED (PERIOD)", "0", CivicCalmTheme.ColorPrimary)
            lblAnalyticsStatApproved = CreateStatCard(pnlCards, 1, "APPROVED / RELEASED", "0", CivicCalmTheme.ColorAccentSG)
            lblAnalyticsStatPending = CreateStatCard(pnlCards, 2, "ACTION REQD / REVISION", "0", CivicCalmTheme.ColorDanger)
            lblAnalyticsStatVelocity = CreateStatCard(pnlCards, 3, "AVG PROCESSING VELOCITY", "0.0 Days", CivicCalmTheme.ColorInfo)

            ' 3. Lower Split Area: Tables (62%) and GAD Demographics (38%)
            Dim tblLower As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 2,
                .RowCount = 1,
                .BackColor = CivicCalmTheme.ColorCanvas
            }
            tblLower.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 62.0F))
            tblLower.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 38.0F))
            tblLower.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

            ' Left Container: Stack of Most Requested & Section Performance
            Dim tblTables As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .RowCount = 2,
                .BackColor = CivicCalmTheme.ColorCanvas,
                .Margin = New Padding(0, 0, 10, 0),
                .MinimumSize = New Size(CInt(Dpi(520.0F)), CInt(Dpi(480.0F)))
            }
            tblTables.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            tblTables.RowStyles.Add(New RowStyle(SizeType.Percent, 50.0F))
            tblTables.RowStyles.Add(New RowStyle(SizeType.Percent, 50.0F))

            ' Table 1: Most Requested Documents Card
            Dim pnlMostReqCard As New Panel With {
                .Dock = DockStyle.Fill,
                .Margin = New Padding(0, 0, 0, 8),
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(12),
                .MinimumSize = New Size(CInt(Dpi(480.0F)), CInt(Dpi(220.0F)))
            }
            ApplyCardBorder(pnlMostReqCard)

            Dim lblMostReqTitle As New Label With {
                .Text = "MOST REQUESTED DOCUMENTS (VOLUME & CLASSIFICATION RANKING)",
                .Font = CivicCalmTheme.FontSectionHeader,
                .ForeColor = CivicCalmTheme.ColorInk,
                .Dock = DockStyle.Top,
                .Height = 24
            }

            dgvMostRequested = New DataGridView With {
                .Dock = DockStyle.Top,
                .Height = CInt(Dpi(126.0F)),
                .ReadOnly = True,
                .AllowUserToAddRows = False,
                .ScrollBars = ScrollBars.Vertical
            }
            DataGridStyler.ApplyCivicStyle(dgvMostRequested)
            lblMostRequestedWatermark = CreateGridWatermark(pnlMostReqCard)

            pnlMostReqChart = New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(0, 6, 0, 0)
            }
            AddHandler pnlMostReqChart.Paint, AddressOf OnPaintCategoryChart
            AddHandler pnlMostReqChart.Resize, Sub() pnlMostReqChart.Invalidate()

            pnlMostReqCard.Controls.Add(pnlMostReqChart)
            pnlMostReqCard.Controls.Add(dgvMostRequested)
            pnlMostReqCard.Controls.Add(lblMostReqTitle)
            tblTables.Controls.Add(pnlMostReqCard, 0, 0)

            ' Table 2: Section Desk Performance Card
            Dim pnlSecCard As New Panel With {
                .Dock = DockStyle.Fill,
                .Margin = New Padding(0),
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(12),
                .MinimumSize = New Size(CInt(Dpi(480.0F)), CInt(Dpi(260.0F)))
            }
            ApplyCardBorder(pnlSecCard)

            Dim lblSecTitle As New Label With {
                .Text = "SECTION DESK ROUTING & OPERATIONAL PERFORMANCE",
                .Font = CivicCalmTheme.FontSectionHeader,
                .ForeColor = CivicCalmTheme.ColorInk,
                .Dock = DockStyle.Top,
                .Height = 24
            }

            dgvSectionBreakdown = New DataGridView With {
                .Dock = DockStyle.Top,
                .Height = CInt(Dpi(150.0F)),
                .ReadOnly = True,
                .AllowUserToAddRows = False,
                .ScrollBars = ScrollBars.Vertical
            }
            DataGridStyler.ApplyCivicStyle(dgvSectionBreakdown)
            lblSectionWatermark = CreateGridWatermark(pnlSecCard)

            pnlSectionChart = New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(0, 6, 0, 0)
            }
            AddHandler pnlSectionChart.Paint, AddressOf OnPaintSectionChart
            AddHandler pnlSectionChart.Resize, Sub() pnlSectionChart.Invalidate()

            pnlSecCard.Controls.Add(pnlSectionChart)
            pnlSecCard.Controls.Add(dgvSectionBreakdown)
            pnlSecCard.Controls.Add(lblSecTitle)
            tblTables.Controls.Add(pnlSecCard, 0, 1)

            ' Right Container: GAD Demographics Card
            pnlGadDemographics = New Panel With {
                .Dock = DockStyle.Fill,
                .Margin = New Padding(0),
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(16),
                .MinimumSize = New Size(CInt(Dpi(300.0F)), CInt(Dpi(480.0F)))
            }
            ApplyCardBorder(pnlGadDemographics)

            Dim pnlGadTopBar As New Panel With {
                .Dock = DockStyle.Top,
                .Height = 4,
                .BackColor = CivicCalmTheme.ColorPrimary
            }

            Dim lblGadHeader As New Label With {
                .Text = "GENDER & DEVELOPMENT (GAD) DEMOGRAPHICS",
                .Font = CivicCalmTheme.FontSectionHeader,
                .ForeColor = CivicCalmTheme.ColorInk,
                .Dock = DockStyle.Top,
                .Height = 26,
                .Padding = New Padding(0, 6, 0, 0)
            }

            Dim lblGadSub As New Label With {
                .Text = "Philippine Magna Carta of Women (RA 9710) Disaggregated Statistics",
                .Font = CivicCalmTheme.FontMicrocopy,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .Dock = DockStyle.Top,
                .Height = 20
            }

            lblGadTotalSubmissions = New Label With {
                .Text = "Total Requesters Analyzed: 0",
                .Font = CivicCalmTheme.FontFieldLabel,
                .ForeColor = CivicCalmTheme.ColorInk,
                .Dock = DockStyle.Top,
                .Height = 26,
                .Padding = New Padding(0, 6, 0, 0)
            }

            ' Ratio bar container
            pnlGadRatioBar = New Panel With {
                .Dock = DockStyle.Top,
                .Height = 28,
                .Margin = New Padding(0, 12, 0, 12),
                .BackColor = CivicCalmTheme.ColorWell
            }
            AddHandler pnlGadRatioBar.Paint, AddressOf OnPaintGadRatioBar
            AddHandler pnlGadRatioBar.Resize, Sub() pnlGadRatioBar.Invalidate()

            lblGadFemaleCount = New Label With {
                .Text = "• Female: 0 (0.0%)",
                .Font = CivicCalmTheme.FontBody,
                .ForeColor = ColorTranslator.FromHtml("#0284C7"),
                .Dock = DockStyle.Top,
                .Height = 24
            }

            lblGadMaleCount = New Label With {
                .Text = "• Male: 0 (0.0%)",
                .Font = CivicCalmTheme.FontBody,
                .ForeColor = ColorTranslator.FromHtml("#0F2A4A"),
                .Dock = DockStyle.Top,
                .Height = 24
            }

            lblGadUndisclosedCount = New Label With {
                .Text = "• Prefer not to say: 0 (0.0%)",
                .Font = CivicCalmTheme.FontBody,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .Dock = DockStyle.Top,
                .Height = 24
            }

            Dim lblGadNotice As New Label With {
                .Text = "Data is automatically synchronized from external portal submissions and internal transmittals to track gender inclusivity under BTA GAD institutional indicators.",
                .Font = CivicCalmTheme.FontMicrocopy,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .Dock = DockStyle.Top,
                .Height = 52,
                .Padding = New Padding(0, 10, 0, 0)
            }

            pnlGadDemographics.Controls.AddRange(New Control() {
                lblGadNotice,
                lblGadUndisclosedCount,
                lblGadMaleCount,
                lblGadFemaleCount,
                pnlGadRatioBar,
                lblGadTotalSubmissions,
                lblGadSub,
                lblGadHeader,
                pnlGadTopBar
            })

            tblLower.Controls.Add(tblTables, 0, 0)
            tblLower.Controls.Add(pnlGadDemographics, 1, 0)

            tblMain.Controls.Add(pnlFilterCard, 0, 0)
            tblMain.Controls.Add(pnlCards, 0, 1)
            tblMain.Controls.Add(tblLower, 0, 2)

            viewAnalytics.Controls.Add(tblMain)
            UiBuffering.EnableDeep(viewAnalytics)
        End Sub

        Private Function GetCategoryColor(cat As String) As Color
            If cat.IndexOf("COMM", StringComparison.OrdinalIgnoreCase) >= 0 OrElse cat.IndexOf("Communication", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return ColorTranslator.FromHtml("#146A3D") ' Primary Green
            ElseIf cat.IndexOf("LEG", StringComparison.OrdinalIgnoreCase) >= 0 OrElse cat.IndexOf("Legislative", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return ColorTranslator.FromHtml("#0369A1") ' Civic Blue
            ElseIf cat.IndexOf("FIN", StringComparison.OrdinalIgnoreCase) >= 0 OrElse cat.IndexOf("Finance", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return ColorTranslator.FromHtml("#8C6414") ' SG Gold / Amber
            ElseIf cat.IndexOf("TO", StringComparison.OrdinalIgnoreCase) >= 0 OrElse cat.IndexOf("Travel", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return ColorTranslator.FromHtml("#0F2A4A") ' Navy
            Else
                Return ColorTranslator.FromHtml("#55606A") ' Slate
            End If
        End Function

        Private Sub OnPaintCategoryChart(sender As Object, e As PaintEventArgs)
            Dim rect = pnlMostReqChart.ClientRectangle
            If rect.Width < 60 OrElse rect.Height < 40 Then Return

            Dim g = e.Graphics
            g.SmoothingMode = SmoothingMode.AntiAlias
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit

            Dim titleRect As New Rectangle(4, 2, rect.Width - 8, 16)
            TextRenderer.DrawText(g, "VOLUME DISTRIBUTION (GRAPHICAL VIEW)", CivicCalmTheme.FontMicrocopy, titleRect, CivicCalmTheme.ColorInkMuted, TextFormatFlags.Left Or TextFormatFlags.VerticalCenter)

            If _categoryChartItems.Count = 0 Then
                Dim emptyRect As New Rectangle(0, 18, rect.Width, rect.Height - 18)
                TextRenderer.DrawText(g, "No volume data to chart.", CivicCalmTheme.FontMicrocopy, emptyRect, CivicCalmTheme.ColorInkMuted, TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter)
                Return
            End If

            Dim topOffset = 20
            Dim availableHeight = rect.Height - topOffset - 4
            If availableHeight < 16 Then Return

            Dim itemCount = Math.Min(_categoryChartItems.Count, 5)
            If itemCount = 0 Then Return
            Dim rowHeight = Math.Max(14, Math.Min(22, availableHeight \ itemCount))

            Dim maxVal = 1
            For Each item In _categoryChartItems
                If item.Count > maxVal Then maxVal = item.Count
            Next

            Dim labelWidth = Math.Min(CInt(Dpi(150.0F)), Math.Max(70, rect.Width \ 3))
            Dim valueWidth = Math.Min(CInt(Dpi(90.0F)), Math.Max(50, rect.Width \ 4))
            Dim chartWidth = Math.Max(20, rect.Width - labelWidth - valueWidth - 16)

            For i As Integer = 0 To itemCount - 1
                Dim item = _categoryChartItems(i)
                Dim y = topOffset + (i * rowHeight)
                If y + rowHeight > rect.Height Then Exit For
                Dim barH = Math.Max(6, rowHeight - 6)
                Dim barY = y + (rowHeight - barH) \ 2

                ' Category label
                Dim lblRect As New Rectangle(4, y, labelWidth - 8, rowHeight)
                Dim truncatedName = item.Category
                If truncatedName.Length > 22 Then truncatedName = truncatedName.Substring(0, 20) & "..."
                TextRenderer.DrawText(g, truncatedName, CivicCalmTheme.FontBody, lblRect, CivicCalmTheme.ColorInk, TextFormatFlags.Left Or TextFormatFlags.VerticalCenter Or TextFormatFlags.EndEllipsis)

                ' Background track
                Dim trackRect As New Rectangle(labelWidth, barY, chartWidth, barH)
                Using bTrack As New SolidBrush(CivicCalmTheme.ColorWell)
                    g.FillRectangle(bTrack, trackRect)
                End Using

                ' Filled bar
                Dim barW = CInt((CDbl(item.Count) / CDbl(maxVal)) * chartWidth)
                If barW > 0 Then
                    Using bBar As New SolidBrush(item.Color)
                        g.FillRectangle(bBar, New Rectangle(labelWidth, barY, barW, barH))
                    End Using
                End If

                ' Border on track
                Using pTrack As New Pen(CivicCalmTheme.ColorBorder)
                    g.DrawRectangle(pTrack, trackRect)
                End Using

                ' Value label
                Dim valRect As New Rectangle(labelWidth + chartWidth + 8, y, valueWidth, rowHeight)
                Dim valStr = $"{item.Count} ({item.Percentage:F1}%)"
                TextRenderer.DrawText(g, valStr, CivicCalmTheme.FontTabular, valRect, CivicCalmTheme.ColorInk, TextFormatFlags.Left Or TextFormatFlags.VerticalCenter)
            Next
        End Sub

        Private Sub OnPaintSectionChart(sender As Object, e As PaintEventArgs)
            Dim rect = pnlSectionChart.ClientRectangle
            If rect.Width < 60 OrElse rect.Height < 40 Then Return

            Dim g = e.Graphics
            g.SmoothingMode = SmoothingMode.AntiAlias
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit

            ' Header
            Dim titleRect As New Rectangle(4, 2, CInt(Dpi(220.0F)), 16)
            TextRenderer.DrawText(g, "SECTION WORKLOAD COMPARISON", CivicCalmTheme.FontMicrocopy, titleRect, CivicCalmTheme.ColorInkMuted, TextFormatFlags.Left Or TextFormatFlags.VerticalCenter)

            ' Legend
            Dim legendX = rect.Width - CInt(Dpi(200.0F))
            If legendX > CInt(Dpi(225.0F)) Then
                Dim legY = 3
                ' Released
                Using bRel As New SolidBrush(CivicCalmTheme.ColorPrimary)
                    g.FillRectangle(bRel, legendX, legY + 2, 8, 8)
                End Using
                TextRenderer.DrawText(g, "Released", CivicCalmTheme.FontMicrocopy, New Point(legendX + 11, legY), CivicCalmTheme.ColorInkMuted)

                ' In Review
                Using bRev As New SolidBrush(CivicCalmTheme.ColorInfo)
                    g.FillRectangle(bRev, legendX + 68, legY + 2, 8, 8)
                End Using
                TextRenderer.DrawText(g, "Review", CivicCalmTheme.FontMicrocopy, New Point(legendX + 79, legY), CivicCalmTheme.ColorInkMuted)

                ' Revision
                Using bAct As New SolidBrush(CivicCalmTheme.ColorDanger)
                    g.FillRectangle(bAct, legendX + 130, legY + 2, 8, 8)
                End Using
                TextRenderer.DrawText(g, "Revision", CivicCalmTheme.FontMicrocopy, New Point(legendX + 141, legY), CivicCalmTheme.ColorInkMuted)
            End If

            If _sectionChartItems.Count = 0 Then
                Dim emptyRect As New Rectangle(0, 18, rect.Width, rect.Height - 18)
                TextRenderer.DrawText(g, "No section workload data to chart.", CivicCalmTheme.FontMicrocopy, emptyRect, CivicCalmTheme.ColorInkMuted, TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter)
                Return
            End If

            Dim topOffset = 20
            Dim availableHeight = rect.Height - topOffset - 4
            If availableHeight < 16 Then Return

            Dim itemCount = Math.Min(_sectionChartItems.Count, 5)
            If itemCount = 0 Then Return
            Dim rowHeight = Math.Max(14, Math.Min(22, availableHeight \ itemCount))

            Dim maxTotal = 1
            For Each item In _sectionChartItems
                If item.TotalAssigned > maxTotal Then maxTotal = item.TotalAssigned
            Next

            Dim labelWidth = Math.Min(CInt(Dpi(130.0F)), Math.Max(70, rect.Width \ 3))
            Dim valueWidth = Math.Min(CInt(Dpi(75.0F)), Math.Max(50, rect.Width \ 4))
            Dim chartWidth = Math.Max(20, rect.Width - labelWidth - valueWidth - 16)

            For i As Integer = 0 To itemCount - 1
                Dim item = _sectionChartItems(i)
                Dim y = topOffset + (i * rowHeight)
                If y + rowHeight > rect.Height Then Exit For
                Dim barH = Math.Max(6, rowHeight - 6)
                Dim barY = y + (rowHeight - barH) \ 2

                ' Section name
                Dim lblRect As New Rectangle(4, y, labelWidth - 8, rowHeight)
                Dim shortName = item.SectionName.Replace(" Section", "")
                TextRenderer.DrawText(g, shortName, CivicCalmTheme.FontBody, lblRect, CivicCalmTheme.ColorInk, TextFormatFlags.Left Or TextFormatFlags.VerticalCenter Or TextFormatFlags.EndEllipsis)

                ' Background track
                Dim trackRect As New Rectangle(labelWidth, barY, chartWidth, barH)
                Using bTrack As New SolidBrush(CivicCalmTheme.ColorWell)
                    g.FillRectangle(bTrack, trackRect)
                End Using

                ' Stacked Bar: Released, InReview, ForRevision
                If item.TotalAssigned > 0 Then
                    Dim totalRatio = CDbl(item.TotalAssigned) / CDbl(maxTotal)
                    Dim totalBarW = CInt(totalRatio * chartWidth)

                    Dim wRel = CInt((CDbl(item.ApprovedReleased) / CDbl(item.TotalAssigned)) * totalBarW)
                    Dim wRev = CInt((CDbl(item.InReview) / CDbl(item.TotalAssigned)) * totalBarW)
                    Dim wAct = totalBarW - wRel - wRev

                    Dim curX = labelWidth
                    If wRel > 0 Then
                        Using bRel As New SolidBrush(CivicCalmTheme.ColorPrimary)
                            g.FillRectangle(bRel, New Rectangle(curX, barY, wRel, barH))
                        End Using
                        curX += wRel
                    End If
                    If wRev > 0 Then
                        Using bRev As New SolidBrush(CivicCalmTheme.ColorInfo)
                            g.FillRectangle(bRev, New Rectangle(curX, barY, wRev, barH))
                        End Using
                        curX += wRev
                    End If
                    If wAct > 0 Then
                        Using bAct As New SolidBrush(CivicCalmTheme.ColorDanger)
                            g.FillRectangle(bAct, New Rectangle(curX, barY, wAct, barH))
                        End Using
                    End If
                End If

                Using pTrack As New Pen(CivicCalmTheme.ColorBorder)
                    g.DrawRectangle(pTrack, trackRect)
                End Using

                ' Value label
                Dim valRect As New Rectangle(labelWidth + chartWidth + 8, y, valueWidth, rowHeight)
                TextRenderer.DrawText(g, $"{item.TotalAssigned} docs", CivicCalmTheme.FontTabular, valRect, CivicCalmTheme.ColorInk, TextFormatFlags.Left Or TextFormatFlags.VerticalCenter)
            Next
        End Sub

        Private Sub OnPaintGadRatioBar(sender As Object, e As PaintEventArgs)
            Dim rect = pnlGadRatioBar.ClientRectangle
            If rect.Width <= 10 OrElse rect.Height <= 4 Then Return

            Dim g = e.Graphics
            g.SmoothingMode = SmoothingMode.AntiAlias

            Dim femaleWidth As Integer = CInt(rect.Width * _gadFemaleRatio)
            Dim maleWidth As Integer = CInt(rect.Width * _gadMaleRatio)
            Dim otherWidth As Integer = rect.Width - femaleWidth - maleWidth

            Using bBg As New SolidBrush(CivicCalmTheme.ColorWell)
                g.FillRectangle(bBg, rect)
            End Using

            Dim curX As Integer = 0
            If femaleWidth > 0 Then
                Using bFemale As New SolidBrush(ColorTranslator.FromHtml("#0284C7"))
                    g.FillRectangle(bFemale, New Rectangle(curX, 0, femaleWidth, rect.Height))
                End Using
                curX += femaleWidth
            End If

            If maleWidth > 0 Then
                Using bMale As New SolidBrush(ColorTranslator.FromHtml("#0F2A4A"))
                    g.FillRectangle(bMale, New Rectangle(curX, 0, maleWidth, rect.Height))
                End Using
                curX += maleWidth
            End If

            If otherWidth > 0 Then
                Using bOther As New SolidBrush(ColorTranslator.FromHtml("#94A3B8"))
                    g.FillRectangle(bOther, New Rectangle(curX, 0, otherWidth, rect.Height))
                End Using
            End If

            Using pBorder As New Pen(CivicCalmTheme.ColorBorder)
                g.DrawRectangle(pBorder, 0, 0, rect.Width - 1, rect.Height - 1)
            End Using
        End Sub

        Public Sub RefreshAnalyticsView()
            If viewAnalytics Is Nothing OrElse Me.IsDisposed OrElse Me.WindowState = FormWindowState.Minimized Then Return

            Dim dtDocs = EmbeddedDB.DataSet.Tables("Documents")
            If dtDocs Is Nothing Then Return

            Dim thresholdDate = ResolveReportingPeriod()
            Dim results = CollectAnalytics(dtDocs, thresholdDate)

            UpdateAnalyticsKpiCards(results)
            PopulateMostRequestedTable(results)
            PopulateSectionBreakdownTable(results)
            UpdateGadDemographics(results)

            lblStatusCount.Text = results.MatchedRows.Count.ToString("N0") & " Analyzed Records"
        End Sub

        ''' <summary>
        ''' Reads the timeframe combo into the start of the reporting window and mirrors the
        ''' human-readable period into the notice label beside it.
        ''' </summary>
        Private Function ResolveReportingPeriod() As DateTime
            Dim tfText = If(cmbAnalyticsTimeframe.SelectedItem, "This Month (30 Days)").ToString()
            Dim thresholdDate As DateTime = DateTime.MinValue
            Select Case tfText
                Case "Today (Day)"
                    thresholdDate = DateTime.Today
                    lblAnalyticsPeriodNotice.Text = "Analytics Period: Today (" & DateTime.Today.ToString("MMM dd, yyyy") & ")"
                Case "This Week (7 Days)"
                    thresholdDate = DateTime.Today.AddDays(-7)
                    lblAnalyticsPeriodNotice.Text = "Analytics Period: Past 7 Days (" & thresholdDate.ToString("MMM dd") & " to " & DateTime.Today.ToString("MMM dd, yyyy") & ")"
                Case "This Month (30 Days)"
                    thresholdDate = New DateTime(DateTime.Today.Year, DateTime.Today.Month, 1)
                    lblAnalyticsPeriodNotice.Text = "Analytics Period: This Month (" & DateTime.Today.ToString("MMMM yyyy") & ")"
                Case "This Year (Year)"
                    thresholdDate = New DateTime(DateTime.Today.Year, 1, 1)
                    lblAnalyticsPeriodNotice.Text = "Analytics Period: Year " & DateTime.Today.Year.ToString()
                Case Else
                    thresholdDate = DateTime.MinValue
                    lblAnalyticsPeriodNotice.Text = "Analytics Period: All Time Archive"
            End Select
            Return thresholdDate
        End Function

        Private Function CollectAnalytics(dtDocs As DataTable, thresholdDate As DateTime) As AnalyticsSummary
            Dim selSection = If(cmbAnalyticsSection.SelectedItem, "All Sections").ToString()
            Dim results As New AnalyticsSummary()

            Dim knownSections = New String() {"Records Section", "Secretariat", "Legislative Section", "Finance Section", "Travel Section"}
            For Each sec In knownSections
                results.SectionStats(sec) = New SectionDeskStats With {.SectionName = sec}
            Next

            Dim totalVelocityDays As Double = 0.0
            Dim velocityCount As Integer = 0

            For Each row As DataRow In dtDocs.Rows
                Dim recDate = ParseRecordedDate(row("DateReceived").ToString())

                If thresholdDate <> DateTime.MinValue AndAlso recDate < thresholdDate Then
                    Continue For
                End If

                Dim sec = RowAssignedSection(row)
                If Not selSection.Equals("All Sections", StringComparison.OrdinalIgnoreCase) AndAlso Not sec.Equals(selSection, StringComparison.OrdinalIgnoreCase) Then
                    Continue For
                End If

                results.MatchedRows.Add(row)

                Dim st = row("CurrentStatus").ToString().ToUpperInvariant()
                If st.Contains("APPROVED") OrElse st.Contains("RELEASED") Then
                    results.TotalApproved += 1
                ElseIf st.Contains("REVISION") OrElse st.Contains("REJECT") Then
                    results.TotalPending += 1
                End If

                ' Processing velocity (days between received and now) is capped at 90 so one
                ' ancient mirrored row cannot dominate the average.
                Dim elapsed = (DateTime.Now - recDate).TotalDays
                If elapsed >= 0 Then
                    totalVelocityDays += Math.Min(elapsed, 90.0)
                    velocityCount += 1
                End If

                Dim docType = row("DocType").ToString()
                If String.IsNullOrWhiteSpace(docType) Then docType = "General Communication"
                If Not results.CategoryCounts.ContainsKey(docType) Then
                    results.CategoryCounts(docType) = 0
                    results.CategoryLatest(docType) = recDate
                End If
                results.CategoryCounts(docType) += 1
                If recDate > results.CategoryLatest(docType) Then results.CategoryLatest(docType) = recDate

                If Not results.SectionStats.ContainsKey(sec) Then
                    results.SectionStats(sec) = New SectionDeskStats With {.SectionName = sec}
                End If
                Dim statObj = results.SectionStats(sec)
                statObj.TotalAssigned += 1
                If st.Contains("REVIEW") Then statObj.InReview += 1
                If st.Contains("REVISION") Then statObj.ForRevision += 1
                If st.Contains("APPROVED") OrElse st.Contains("RELEASED") Then statObj.ApprovedReleased += 1

                Dim gnd = InferRequesterGender(row)
                If gnd.Equals("Female", StringComparison.OrdinalIgnoreCase) Then
                    results.FemaleCount += 1
                ElseIf gnd.Equals("Male", StringComparison.OrdinalIgnoreCase) Then
                    results.MaleCount += 1
                Else
                    results.OtherCount += 1
                End If
            Next

            results.AvgVelocityDays = If(velocityCount > 0, totalVelocityDays / velocityCount, 0.0)
            Return results
        End Function

        ' DateReceived is a free-form string: SQL mirrors write "yyyy-MM-dd HH:mm:ss" (parse
        ' invariant first), the embedded cache may hold any culture format, and unparseable
        ' rows count as received today.
        Private Shared Function ParseRecordedDate(dateStr As String) As DateTime
            Dim recDate As DateTime
            If DateTime.TryParse(dateStr, CultureInfo.InvariantCulture, DateTimeStyles.None, recDate) Then Return recDate
            If DateTime.TryParse(dateStr, recDate) Then Return recDate
            Return DateTime.UtcNow
        End Function

        Private Shared Function RowAssignedSection(row As DataRow) As String
            Dim sec = If(row.Table.Columns.Contains("AssignedSection") AndAlso Not IsDBNull(row("AssignedSection")), row("AssignedSection").ToString(), "Records Section")
            If String.IsNullOrWhiteSpace(sec) Then sec = "Records Section"
            Return sec
        End Function

        ''' <summary>
        ''' The recorded gender, falling back to the "Gender: ..." tag some portal imports
        ''' carry in their remarks. Unrecorded gender counts under "prefer not to say".
        ''' </summary>
        Private Shared Function InferRequesterGender(row As DataRow) As String
            Dim gnd = ""
            If row.Table.Columns.Contains("RequesterGender") AndAlso Not IsDBNull(row("RequesterGender")) Then
                gnd = row("RequesterGender").ToString().Trim()
            End If
            If String.IsNullOrEmpty(gnd) Then
                Dim rm = If(row.Table.Columns.Contains("Remarks") AndAlso Not IsDBNull(row("Remarks")), row("Remarks").ToString(), "")
                If rm.IndexOf("Gender: Female", StringComparison.OrdinalIgnoreCase) >= 0 Then
                    gnd = "Female"
                ElseIf rm.IndexOf("Gender: Male", StringComparison.OrdinalIgnoreCase) >= 0 Then
                    gnd = "Male"
                End If
            End If
            Return gnd
        End Function

        Private Sub UpdateAnalyticsKpiCards(results As AnalyticsSummary)
            lblAnalyticsStatTotal.Text = results.MatchedRows.Count.ToString("N0")
            lblAnalyticsStatApproved.Text = results.TotalApproved.ToString("N0")
            lblAnalyticsStatPending.Text = results.TotalPending.ToString("N0")
            lblAnalyticsStatVelocity.Text = results.AvgVelocityDays.ToString("F1") & " Days"
        End Sub

        Private Sub PopulateMostRequestedTable(results As AnalyticsSummary)
            Dim dtMostReq As New DataTable()
            dtMostReq.Columns.Add("Rank", GetType(Integer))
            dtMostReq.Columns.Add("Category", GetType(String))
            dtMostReq.Columns.Add("RequestCount", GetType(Integer))
            dtMostReq.Columns.Add("SharePercent", GetType(String))
            dtMostReq.Columns.Add("LatestRequest", GetType(String))

            Dim sortedCats = New List(Of KeyValuePair(Of String, Integer))(results.CategoryCounts)
            sortedCats.Sort(Function(a, b) b.Value.CompareTo(a.Value))

            Dim rank As Integer = 1
            Dim totalVolume As Integer = results.MatchedRows.Count
            _categoryChartItems.Clear()

            For Each kvp In sortedCats
                Dim pct = If(totalVolume > 0, (kvp.Value / CDbl(totalVolume)) * 100.0, 0.0)
                Dim lat = If(results.CategoryLatest.ContainsKey(kvp.Key), results.CategoryLatest(kvp.Key).ToString("yyyy-MM-dd"), "-")
                dtMostReq.Rows.Add(rank, kvp.Key, kvp.Value, pct.ToString("F1") & "%", lat)

                _categoryChartItems.Add(New CategoryChartItem With {
                    .Category = kvp.Key,
                    .Count = kvp.Value,
                    .Percentage = pct,
                    .Color = GetCategoryColor(kvp.Key)
                })
                rank += 1
            Next

            dgvMostRequested.DataSource = dtMostReq
            FormatMostRequestedGrid(dgvMostRequested)
            If dtMostReq.Rows.Count = 0 Then
                DataGridStyler.SetEmptyState(dgvMostRequested, lblMostRequestedWatermark, "No document requests recorded in this reporting period.")
            Else
                DataGridStyler.SetPopulatedState(dgvMostRequested, lblMostRequestedWatermark)
            End If

            If pnlMostReqChart IsNot Nothing Then pnlMostReqChart.Invalidate()
        End Sub

        Private Sub PopulateSectionBreakdownTable(results As AnalyticsSummary)
            Dim dtSec As New DataTable()
            dtSec.Columns.Add("SectionName", GetType(String))
            dtSec.Columns.Add("TotalAssigned", GetType(Integer))
            dtSec.Columns.Add("InReview", GetType(Integer))
            dtSec.Columns.Add("ForRevision", GetType(Integer))
            dtSec.Columns.Add("ApprovedReleased", GetType(Integer))

            _sectionChartItems.Clear()
            For Each kvp In results.SectionStats
                Dim s = kvp.Value
                dtSec.Rows.Add(s.SectionName, s.TotalAssigned, s.InReview, s.ForRevision, s.ApprovedReleased)
                _sectionChartItems.Add(s)
            Next

            dgvSectionBreakdown.DataSource = dtSec
            FormatSectionBreakdownGrid(dgvSectionBreakdown)
            If dtSec.Rows.Count = 0 Then
                DataGridStyler.SetEmptyState(dgvSectionBreakdown, lblSectionWatermark, "No section desk activity in this reporting period.")
            Else
                DataGridStyler.SetPopulatedState(dgvSectionBreakdown, lblSectionWatermark)
            End If

            If pnlSectionChart IsNot Nothing Then pnlSectionChart.Invalidate()
        End Sub

        Private Sub UpdateGadDemographics(results As AnalyticsSummary)
            Dim gadTotal = results.FemaleCount + results.MaleCount + results.OtherCount
            lblGadTotalSubmissions.Text = "Total Requesters Analyzed: " & gadTotal.ToString("N0")
            If gadTotal > 0 Then
                Dim fPct = (results.FemaleCount / CDbl(gadTotal)) * 100.0
                Dim mPct = (results.MaleCount / CDbl(gadTotal)) * 100.0
                Dim oPct = (results.OtherCount / CDbl(gadTotal)) * 100.0

                lblGadFemaleCount.Text = $"• Female: {results.FemaleCount} ({fPct:F1}%)"
                lblGadMaleCount.Text = $"• Male: {results.MaleCount} ({mPct:F1}%)"
                lblGadUndisclosedCount.Text = $"• Prefer not to say: {results.OtherCount} ({oPct:F1}%)"

                _gadFemaleRatio = CSng(results.FemaleCount / CDbl(gadTotal))
                _gadMaleRatio = CSng(results.MaleCount / CDbl(gadTotal))
                _gadOtherRatio = CSng(results.OtherCount / CDbl(gadTotal))
            Else
                lblGadFemaleCount.Text = "• Female: 0 (0.0%)"
                lblGadMaleCount.Text = "• Male: 0 (0.0%)"
                lblGadUndisclosedCount.Text = "• Prefer not to say: 0 (0.0%)"
                _gadFemaleRatio = 0.0F
                _gadMaleRatio = 0.0F
                _gadOtherRatio = 0.0F
            End If
            pnlGadRatioBar.Invalidate()
        End Sub

        Private Sub FormatMostRequestedGrid(dgv As DataGridView)
            If dgv Is Nothing OrElse dgv.Columns.Count = 0 Then Return
            Try
                dgv.SuspendLayout()
                dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
                dgv.ColumnHeadersHeight = CInt(Dpi(28.0F))
                dgv.RowTemplate.Height = CInt(Dpi(24.0F))
                dgv.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.False
                For Each col As DataGridViewColumn In dgv.Columns
                    Select Case col.Name
                        Case "Rank"
                            col.HeaderText = "#"
                            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                            col.MinimumWidth = 35
                            col.FillWeight = 35
                        Case "Category"
                            col.HeaderText = "Category"
                            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                            col.MinimumWidth = 130
                            col.FillWeight = 200
                        Case "RequestCount"
                            col.HeaderText = "Requests"
                            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                            col.DefaultCellStyle.Font = CivicCalmTheme.FontTabular
                            col.MinimumWidth = 70
                            col.FillWeight = 75
                        Case "SharePercent"
                            col.HeaderText = "Share"
                            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                            col.DefaultCellStyle.Font = CivicCalmTheme.FontTabular
                            col.MinimumWidth = 60
                            col.FillWeight = 65
                        Case "LatestRequest"
                            col.HeaderText = "Latest Ingest"
                            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                            col.MinimumWidth = 85
                            col.FillWeight = 90
                    End Select
                    col.DefaultCellStyle.Padding = New Padding(4, 2, 4, 2)
                    col.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                Next
            Finally
                dgv.ResumeLayout()
            End Try
        End Sub

        Private Sub FormatSectionBreakdownGrid(dgv As DataGridView)
            If dgv Is Nothing OrElse dgv.Columns.Count = 0 Then Return
            Try
                dgv.SuspendLayout()
                dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
                dgv.ColumnHeadersHeight = CInt(Dpi(28.0F))
                dgv.RowTemplate.Height = CInt(Dpi(24.0F))
                dgv.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.False
                For Each col As DataGridViewColumn In dgv.Columns
                    Select Case col.Name
                        Case "SectionName"
                            col.HeaderText = "Section Desk"
                            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                            col.MinimumWidth = 120
                            col.FillWeight = 180
                        Case "TotalAssigned"
                            col.HeaderText = "Assigned"
                            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                            col.DefaultCellStyle.Font = CivicCalmTheme.FontTabular
                            col.MinimumWidth = 65
                            col.FillWeight = 70
                        Case "InReview"
                            col.HeaderText = "In Review"
                            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                            col.DefaultCellStyle.Font = CivicCalmTheme.FontTabular
                            col.MinimumWidth = 70
                            col.FillWeight = 75
                        Case "ForRevision"
                            col.HeaderText = "Revision"
                            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                            col.DefaultCellStyle.Font = CivicCalmTheme.FontTabular
                            col.MinimumWidth = 65
                            col.FillWeight = 70
                        Case "ApprovedReleased"
                            col.HeaderText = "Released"
                            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                            col.DefaultCellStyle.Font = CivicCalmTheme.FontTabular
                            col.MinimumWidth = 65
                            col.FillWeight = 70
                    End Select
                    col.DefaultCellStyle.Padding = New Padding(4, 2, 4, 2)
                    col.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                Next
            Finally
                dgv.ResumeLayout()
            End Try
        End Sub

        Private Class CategoryChartItem
            Public Property Category As String = ""
            Public Property Count As Integer = 0
            Public Property Percentage As Double = 0.0
            Public Property Color As Color = CivicCalmTheme.ColorPrimary
        End Class

        Private Class SectionDeskStats
            Public Property SectionName As String = ""
            Public Property TotalAssigned As Integer = 0
            Public Property InReview As Integer = 0
            Public Property ForRevision As Integer = 0
            Public Property ApprovedReleased As Integer = 0
        End Class

        ''' <summary>
        ''' One pass over the Documents table: every tally the analytics cards, charts, and
        ''' the GAD panel render. Produced by CollectAnalytics, consumed by the render methods.
        ''' </summary>
        Private Class AnalyticsSummary
            Public ReadOnly MatchedRows As New List(Of DataRow)()
            Public Property TotalApproved As Integer = 0
            Public Property TotalPending As Integer = 0
            Public Property AvgVelocityDays As Double = 0.0
            Public ReadOnly CategoryCounts As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
            Public ReadOnly CategoryLatest As New Dictionary(Of String, DateTime)(StringComparer.OrdinalIgnoreCase)
            Public ReadOnly SectionStats As New Dictionary(Of String, SectionDeskStats)(StringComparer.OrdinalIgnoreCase)
            Public Property FemaleCount As Integer = 0
            Public Property MaleCount As Integer = 0
            Public Property OtherCount As Integer = 0
        End Class
    End Class
End Namespace
