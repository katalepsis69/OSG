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

            lblStatTotalDocs = CreateStatCard(pnlCards, 0, "TOTAL OSG DOCUMENTS", "0", CivicCalmTheme.ColorPrimary)
            lblStatDirectives = CreateStatCard(pnlCards, 1, "PENDING REVIEW", "0", CivicCalmTheme.ColorInfo)
            lblStatActiveRoute = CreateStatCard(pnlCards, 2, "ACTION REQD (REVISION)", "0", CivicCalmTheme.ColorDanger)
            lblStatVaultStorage = CreateStatCard(pnlCards, 3, "APPROVED / RELEASED", "0", CivicCalmTheme.ColorAccentSG)

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
                .Height = 36,
                .FlowDirection = FlowDirection.LeftToRight,
                .WrapContents = False,
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

            btnDashResetFilters = New Button With {
                .Text = " &Reset Filters",
                .Size = New Size(110, 28),
                .Font = CivicCalmTheme.FontMicrocopy,
                .BackColor = CivicCalmTheme.ColorWell,
                .ForeColor = CivicCalmTheme.ColorInk,
                .FlatStyle = FlatStyle.Flat,
                .Margin = New Padding(16, 1, 0, 0),
                .Image = AppAssets.GetIcon("arrow-clockwise", 14, CivicCalmTheme.ColorInk),
                .ImageAlign = ContentAlignment.MiddleLeft,
                .TextAlign = ContentAlignment.MiddleCenter,
                .TextImageRelation = TextImageRelation.ImageBeforeText,
                .Padding = New Padding(4, 0, 4, 0)
            }
            btnDashResetFilters.FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder
            AddHandler btnDashResetFilters.Click, Sub()
                                                      cmbDashSection.SelectedIndex = 0
                                                      cmbDashCategory.SelectedIndex = 0
                                                  End Sub

            Dim btnPrintDashSlip As New Button With {
                .Text = " &Print Slip",
                .Size = New Size(105, 28),
                .Font = CivicCalmTheme.FontMicrocopy,
                .BackColor = CivicCalmTheme.ColorWell,
                .ForeColor = CivicCalmTheme.ColorInk,
                .FlatStyle = FlatStyle.Flat,
                .Margin = New Padding(8, 1, 0, 0),
                .Image = AppAssets.GetIcon("printer", 14, CivicCalmTheme.ColorInk),
                .ImageAlign = ContentAlignment.MiddleLeft,
                .TextAlign = ContentAlignment.MiddleCenter,
                .TextImageRelation = TextImageRelation.ImageBeforeText,
                .Padding = New Padding(4, 0, 4, 0)
            }
            btnPrintDashSlip.FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder
            AddHandler btnPrintDashSlip.Click, Sub() PrintSelectedDocumentRoutingSlip(dgvDashRecent)

            pnlFilterBar.Controls.AddRange(New Control() {lblSecFilter, cmbDashSection, lblCatFilter, cmbDashCategory, btnDashResetFilters, btnPrintDashSlip})

            dgvDashRecent = New DataGridView With {
                .Dock = DockStyle.Fill,
                .ReadOnly = True,
                .AllowUserToAddRows = False
            }
            DataGridStyler.ApplyCivicStyle(dgvDashRecent)
            AddHandler dgvDashRecent.CellDoubleClick, Sub(s, e) If e.RowIndex >= 0 Then OpenSelectedDocumentDetail(dgvDashRecent)

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
