Option Explicit On
Option Strict On

Imports System
Imports System.Data
Imports System.Drawing
Imports System.Windows.Forms

Namespace BTA_OSG
    Partial Public Class FormMain
        Private Sub SetupDashboardView()
            viewDashboard = New Panel With {.Dock = DockStyle.Fill, .BackColor = CivicCalmTheme.ColorCanvas}

            Dim tblDashboard As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .RowCount = 2,
                .BackColor = CivicCalmTheme.ColorCanvas,
                .Padding = New Padding(0)
            }
            tblDashboard.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            tblDashboard.RowStyles.Add(New RowStyle(SizeType.Absolute, 96.0F))
            tblDashboard.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

            Dim pnlCards As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .Margin = New Padding(0, 0, 0, 12),
                .ColumnCount = 4,
                .RowCount = 1,
                .BackColor = CivicCalmTheme.ColorCanvas
            }
            pnlCards.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 25.0F))
            pnlCards.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 25.0F))
            pnlCards.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 25.0F))
            pnlCards.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 25.0F))
            pnlCards.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

            ' The four cards are office-wide counts; only the grid below follows the section
            ' and category filters, so the labels state their scope.
            lblStatTotalDocs = CreateStatCard(pnlCards, 0, "TOTAL OSG DOCUMENTS (ALL)", "0", CivicCalmTheme.ColorPrimary)
            lblStatForReview = CreateStatCard(pnlCards, 1, "PENDING REVIEW (ALL)", "0", CivicCalmTheme.ColorInfo)
            lblStatForRevision = CreateStatCard(pnlCards, 2, "ACTION REQD (REVISION, ALL)", "0", CivicCalmTheme.ColorDanger)
            lblStatVaultStorage = CreateStatCard(pnlCards, 3, "APPROVED / RELEASED (ALL)", "0", CivicCalmTheme.ColorAccentSG)

            Dim pnlGridCard As New Panel With {
                .Dock = DockStyle.Fill,
                .Margin = New Padding(0),
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(16, 12, 16, 12)
            }
            ApplyCardBorder(pnlGridCard)

            Dim lblRecHeader As New Label With {
                .Text = "OSG DOCUMENT STATUS & MONITORING DESK",
                .UseMnemonic = False,
                .Font = CivicCalmTheme.FontSectionHeader,
                .ForeColor = CivicCalmTheme.ColorInk,
                .Dock = DockStyle.Top,
                .Height = 28
            }

            Dim pnlFilterBar As New FlowLayoutPanel With {
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .FlowDirection = FlowDirection.LeftToRight,
                .WrapContents = True,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(0, 2, 0, 4)
            }

            Dim lblSecFilter As New Label With {
                .Text = "Section Desk:",
                .UseMnemonic = False,
                .Font = CivicCalmTheme.FontFieldLabel,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .AutoSize = True,
                .Margin = New Padding(0, 6, 4, 0)
            }
            cmbDashSection = New ComboBox With {
                .DropDownStyle = ComboBoxStyle.DropDownList,
                .Width = CInt(Dpi(180.0F)),
                .Height = 28,
                .Font = CivicCalmTheme.FontBody,
                .BackColor = CivicCalmTheme.ColorSurface,
                .ForeColor = CivicCalmTheme.ColorInk
            }
            cmbDashSection.Items.AddRange(New Object() {"All Sections", "Records Section", "Secretariat", "Legislative Section", "Finance Section", "Travel Section"})
            cmbDashSection.SelectedIndex = 0
            AddHandler cmbDashSection.SelectedIndexChanged, Sub() RefreshActiveTabGrid()

            Dim lblCatFilter As New Label With {
                .Text = "Data Bank:",
                .UseMnemonic = False,
                .Font = CivicCalmTheme.FontFieldLabel,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .AutoSize = True,
                .Margin = New Padding(16, 6, 4, 0)
            }
            cmbDashCategory = New ComboBox With {
                .DropDownStyle = ComboBoxStyle.DropDownList,
                .Width = CInt(Dpi(180.0F)),
                .Height = 28,
                .Font = CivicCalmTheme.FontBody,
                .BackColor = CivicCalmTheme.ColorSurface,
                .ForeColor = CivicCalmTheme.ColorInk
            }
            cmbDashCategory.Items.AddRange(New Object() {"All Categories", "Regular Communication", "Legislative", "Finance", "Travel Order"})
            cmbDashCategory.SelectedIndex = 0
            AddHandler cmbDashCategory.SelectedIndexChanged, Sub() RefreshActiveTabGrid()

            Dim lblStatusFilter As New Label With {
                .Text = "Status:",
                .UseMnemonic = False,
                .Font = CivicCalmTheme.FontFieldLabel,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .AutoSize = True,
                .Margin = New Padding(16, 6, 4, 0)
            }
            cmbDashStatus = New ComboBox With {
                .DropDownStyle = ComboBoxStyle.DropDownList,
                .Width = CInt(Dpi(150.0F)),
                .Height = 28,
                .Font = CivicCalmTheme.FontBody,
                .BackColor = CivicCalmTheme.ColorSurface,
                .ForeColor = CivicCalmTheme.ColorInk
            }
            ' Codes come from DocumentStatus (single source of truth); the grid filter
            ' maps the chosen display name back through the same module.
            cmbDashStatus.Items.Add("All Statuses")
            For Each code In DocumentStatus.FilterableStatusCodes
                cmbDashStatus.Items.Add(DocumentStatus.DisplayName(code))
            Next
            cmbDashStatus.SelectedIndex = 0
            AddHandler cmbDashStatus.SelectedIndexChanged, Sub() RefreshActiveTabGrid()

            chkDashOnlyMine = New CheckBox With {
                .Text = "Assigned to &Me",
                .AutoSize = True,
                .Font = CivicCalmTheme.FontFieldLabel,
                .ForeColor = CivicCalmTheme.ColorInk,
                .Margin = New Padding(16, 7, 0, 0)
            }
            AddHandler chkDashOnlyMine.CheckedChanged, Sub() RefreshActiveTabGrid()

            btnDashResetFilters = CivicCalmTheme.CreateIconButton("&Reset Filters", "arrow-clockwise", 14, margin:=New Padding(16, 1, 0, 0))
            AddHandler btnDashResetFilters.Click, Sub()
                                                      cmbDashSection.SelectedIndex = 0
                                                      cmbDashCategory.SelectedIndex = 0
                                                      cmbDashStatus.SelectedIndex = 0
                                                      chkDashOnlyMine.Checked = False
                                                  End Sub

            Dim btnQuickView = CivicCalmTheme.CreateIconButton("&Quick View", "eye", 14, minHeight:=28, margin:=New Padding(8, 1, 0, 0))
            AddHandler btnQuickView.Click, Sub() OpenSelectedDocumentQuickView(dgvDashRecent)

            Dim btnPrintDashSlip = CivicCalmTheme.CreateIconButton("&Print Slip", "printer", 14, minHeight:=28, margin:=New Padding(8, 1, 0, 0))
            AddHandler btnPrintDashSlip.Click, Sub() PrintSelectedDocumentRoutingSlip(dgvDashRecent)

            btnRouteSelected = CivicCalmTheme.CreateIconButton("Route &Selection", "arrows-split", 14, minHeight:=28, margin:=New Padding(8, 1, 0, 0))
            AddHandler btnRouteSelected.Click, Sub() RouteSelectedDocuments(dgvDashRecent)

            pnlFilterBar.Controls.AddRange(New Control() {lblSecFilter, cmbDashSection, lblCatFilter, cmbDashCategory, lblStatusFilter, cmbDashStatus, chkDashOnlyMine, btnDashResetFilters, btnQuickView, btnPrintDashSlip, btnRouteSelected})

            dgvDashRecent = New DataGridView With {
                .Dock = DockStyle.Fill,
                .ReadOnly = True,
                .AllowUserToAddRows = False
            }
            DataGridStyler.ApplyCivicStyle(dgvDashRecent)
            AddHandler dgvDashRecent.CellDoubleClick, Sub(s, e) If e.RowIndex >= 0 Then OpenSelectedDocumentDetail(dgvDashRecent)
            AddHandler dgvDashRecent.KeyDown, Sub(s, e) HandleDocumentGridKeys(dgvDashRecent, e)

            lblDashWatermark = CreateGridWatermark(pnlGridCard)

            pnlGridCard.Controls.Add(dgvDashRecent)
            pnlGridCard.Controls.Add(pnlFilterBar)
            pnlGridCard.Controls.Add(lblRecHeader)

            tblDashboard.Controls.Add(pnlCards, 0, 0)
            tblDashboard.Controls.Add(pnlGridCard, 0, 1)

            viewDashboard.Controls.Add(tblDashboard)
        End Sub

        Private Function CreateStatCard(parent As TableLayoutPanel, colIndex As Integer, title As String, initVal As String, accentBg As Color) As Label
            Dim pnlCard As New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Margin = New Padding(If(colIndex = 0, 0, 4), 0, If(colIndex = 3, 0, 4), 0)
            }
            ApplyCardBorder(pnlCard)

            Dim pnlBar As New Panel With {
                .Dock = DockStyle.Top,
                .Height = 4,
                .BackColor = accentBg
            }
            Dim lblT As New Label With {
                .Text = title,
                .UseMnemonic = False,
                .Font = CivicCalmTheme.FontMicrocopy,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .Dock = DockStyle.Top,
                .Height = 22,
                .Padding = New Padding(12, 6, 0, 0)
            }
            Dim lblV As New Label With {
                .Text = initVal,
                .UseMnemonic = False,
                .Font = New Font("Segoe UI", 16.0F, FontStyle.Bold),
                .ForeColor = CivicCalmTheme.ColorInk,
                .Dock = DockStyle.Fill,
                .Padding = New Padding(12, 0, 0, 4),
                .TextAlign = ContentAlignment.MiddleLeft
            }

            pnlCard.Controls.Add(lblV)
            pnlCard.Controls.Add(lblT)
            pnlCard.Controls.Add(pnlBar)
            parent.Controls.Add(pnlCard, colIndex, 0)
            Return lblV
        End Function
    End Class
End Namespace
