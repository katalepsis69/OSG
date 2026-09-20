Option Explicit On
Option Strict On

Imports System
Imports System.Data
Imports System.Diagnostics
Imports System.Drawing
Imports System.Windows.Forms

Namespace BTA_OSG
    Public Class FormDocumentDetail
        Inherits Form

        Private DocID As Integer
        Private MainFrm As FormMain
        Private DocRow As DataRow

        Private pnlHeader As Panel
        Private lblCode As Label
        Private lblTitle As Label
        Private lblStatusBadge As Label

        Private tabDetail As TabControl
        Private tabOverview As TabPage
        Private tabDirectives As TabPage
        Private tabRouting As TabPage
        Private tabMovements As TabPage

        Private dgvDirectives As DataGridView
        Private dgvRouting As DataGridView
        Private dgvMovements As DataGridView

        Private pnlOverviewTable As TableLayoutPanel
        Private btnClose As Button

        Public Sub New(id As Integer, parentForm As FormMain)
            DocID = id
            MainFrm = parentForm
            InitializeForm()
        End Sub

        Private Sub InitializeForm()
            If Not LoadDocData() Then Return

            Me.Text = String.Format("Document Details & Specifications : [{0}] {1}", DocRow("DocCode"), DocRow("Title"))
            Me.Size = New Size(1080, 760)
            Me.MinimumSize = New Size(920, 640)
            Me.StartPosition = FormStartPosition.CenterParent
            Me.Font = CivicCalmTheme.FontBody
            Me.BackColor = CivicCalmTheme.ColorCanvas

            ' Top Header Panel
            pnlHeader = New Panel With {
                .Dock = DockStyle.Top,
                .Height = 125,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(20, 16, 20, 16)
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
            tblHeader.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 65.0F))
            tblHeader.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 35.0F))

            Dim pnlTitles As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.TopDown,
                .WrapContents = False,
                .BackColor = CivicCalmTheme.ColorSurface
            }

            lblCode = New Label With {
                .Text = String.Format("DOC CODE: {0}   |   TYPE: {1}", DocRow("DocCode"), DocRow("DocType").ToString().ToUpperInvariant()),
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
                .MaximumSize = New Size(620, 40),
                .Margin = New Padding(0, 0, 0, 4)
            }

            lblStatusBadge = New Label With {
                .Text = String.Format("Status: {0}  |  Assigned Staff: {1}", DocRow("CurrentStatus"), DocRow("AssignedStaff")),
                .Font = CivicCalmTheme.FontFieldLabel,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .AutoSize = True
            }
            pnlTitles.Controls.AddRange(New Control() {lblCode, lblTitle, lblStatusBadge})

            Dim pnlActions As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.RightToLeft,
                .WrapContents = True,
                .BackColor = CivicCalmTheme.ColorSurface
            }

            Dim btnRoute As New Button With {
                .Text = "&Route Document",
                .Size = New Size(150, 34),
                .BackColor = CivicCalmTheme.ColorPrimary,
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(6, 4, 0, 4)
            }
            btnRoute.FlatAppearance.BorderSize = 0
            AddHandler btnRoute.Click, AddressOf OnRouteDocument

            Dim btnMove As New Button With {
                .Text = "&Transfer Storage",
                .Size = New Size(150, 34),
                .BackColor = CivicCalmTheme.ColorWell,
                .ForeColor = CivicCalmTheme.ColorInk,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(6, 4, 0, 4)
            }
            btnMove.FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder
            AddHandler btnMove.Click, AddressOf OnMoveStorage

            Dim btnLaunchPdf As New Button With {
                .Text = "&Launch PDF",
                .Size = New Size(120, 34),
                .BackColor = CivicCalmTheme.ColorWell,
                .ForeColor = CivicCalmTheme.ColorInk,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(6, 4, 0, 4)
            }
            btnLaunchPdf.FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder
            AddHandler btnLaunchPdf.Click, AddressOf OnLaunchPDF

            btnClose = New Button With {
                .Text = "&Close",
                .Size = New Size(90, 34),
                .BackColor = CivicCalmTheme.ColorWell,
                .ForeColor = CivicCalmTheme.ColorInk,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(6, 4, 0, 4)
            }
            btnClose.FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder
            AddHandler btnClose.Click, Sub() Me.Close()

            Me.CancelButton = btnClose

            pnlActions.Controls.AddRange(New Control() {btnClose, btnLaunchPdf, btnMove, btnRoute})

            tblHeader.Controls.Add(pnlTitles, 0, 0)
            tblHeader.Controls.Add(pnlActions, 1, 0)
            pnlHeader.Controls.Add(tblHeader)

            Me.Controls.Add(pnlHeader)

            ' Tab Navigation
            tabDetail = New TabControl With {
                .Dock = DockStyle.Fill,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Padding = New Point(12, 6)
            }

            tabOverview = New TabPage(" Document Overview & Specifications ")
            tabDirectives = New TabPage(" SG Directives Timeline ")
            tabRouting = New TabPage(" Office Routing Logs ")
            tabMovements = New TabPage(" Physical Storage Movement History ")

            tabOverview.BackColor = CivicCalmTheme.ColorSurface
            tabDirectives.BackColor = CivicCalmTheme.ColorSurface
            tabRouting.BackColor = CivicCalmTheme.ColorSurface
            tabMovements.BackColor = CivicCalmTheme.ColorSurface

            SetupOverviewTab()

            dgvDirectives = CreateDetailGrid()
            dgvRouting = CreateDetailGrid()
            dgvMovements = CreateDetailGrid()

            tabDirectives.Controls.Add(dgvDirectives)
            tabRouting.Controls.Add(dgvRouting)
            tabMovements.Controls.Add(dgvMovements)

            tabDetail.TabPages.AddRange(New TabPage() {tabOverview, tabDirectives, tabRouting, tabMovements})
            Me.Controls.Add(tabDetail)
            pnlHeader.SendToBack()

            RefreshGrids()
        End Sub

        Private Function LoadDocData() As Boolean
            Dim rows = EmbeddedDB.DataSet.Tables("Documents").Select("DocumentID = " & DocID)
            If rows.Length = 0 Then
                MessageBox.Show("Document record not found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.Close()
                Return False
            End If
            DocRow = rows(0)
            Return True
        End Function

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

            pnlOverviewTable.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 30.0F))
            pnlOverviewTable.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 70.0F))

            AddOverviewRow("Document System ID:", "#" & DocRow("DocumentID").ToString())
            AddOverviewRow("Document Code:", DocRow("DocCode").ToString())
            AddOverviewRow("Document Classification Type:", DocRow("DocType").ToString())
            AddOverviewRow("Official Document Title:", DocRow("Title").ToString())
            AddOverviewRow("Date Received / Registered:", DocRow("DateReceived").ToString())
            AddOverviewRow("Originating Office:", DocRow("OriginatingOffice").ToString())
            AddOverviewRow("Destination Office:", DocRow("DestinationOffice").ToString())
            AddOverviewRow("Cabinet Landmark ID:", DocRow("CabinetID").ToString())
            AddOverviewRow("Shelf Landmark No:", DocRow("ShelfNo").ToString())
            AddOverviewRow("Box Landmark Code:", DocRow("BoxCode").ToString())
            AddOverviewRow("Full Physical Location String:", String.Format("Cabinet {0} / Shelf {1} / Box {2}", DocRow("CabinetID"), DocRow("ShelfNo"), DocRow("BoxCode")))
            AddOverviewRow("Current Processing Status:", DocRow("CurrentStatus").ToString())
            AddOverviewRow("Assigned OSG Staff Member:", DocRow("AssignedStaff").ToString())
            AddOverviewRow("Google Drive Soft Copy URL:", DocRow("GDriveURL").ToString())

            pnlScroll.Controls.Add(pnlOverviewTable)
            tabOverview.Controls.Add(pnlScroll)
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

        Private Function CreateDetailGrid() As DataGridView
            Dim dgv As New DataGridView With {
                .Dock = DockStyle.Fill,
                .ReadOnly = True,
                .AllowUserToAddRows = False
            }
            DataGridStyler.ApplyCivicStyle(dgv)
            Return dgv
        End Function

        Private Sub RefreshGrids()
            If Not LoadDocData() Then Return

            lblCode.Text = String.Format("DOC CODE: {0}   |   TYPE: {1}", DocRow("DocCode"), DocRow("DocType").ToString().ToUpperInvariant())
            lblTitle.Text = DocRow("Title").ToString()
            lblStatusBadge.Text = String.Format("Status: {0}  |  Assigned Staff: {1}", DocRow("CurrentStatus"), DocRow("AssignedStaff"))

            SetupOverviewTab()

            Dim dvDir As New DataView(EmbeddedDB.DataSet.Tables("Directives"))
            dvDir.RowFilter = "DocumentID = " & DocID
            dgvDirectives.DataSource = dvDir.ToTable()
            DataGridStyler.FormatDirectiveColumns(dgvDirectives)

            Dim dvRoute As New DataView(EmbeddedDB.DataSet.Tables("RoutingLogs"))
            dvRoute.RowFilter = "DocumentID = " & DocID
            dgvRouting.DataSource = dvRoute.ToTable()
            DataGridStyler.FormatRoutingColumns(dgvRouting)

            Dim dvMove As New DataView(EmbeddedDB.DataSet.Tables("Movements"))
            dvMove.RowFilter = "DocumentID = " & DocID
            dgvMovements.DataSource = dvMove.ToTable()
            DataGridStyler.FormatMovementColumns(dgvMovements)
        End Sub

        Private Sub OnRouteDocument(sender As Object, e As EventArgs)
            If MainFrm.CurrentUser Is Nothing Then
                MessageBox.Show("Authentication Required.", "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Using dlg As New FormRouteDocument(DocID, DocRow("DestinationOffice").ToString(), MainFrm.CurrentUser("FullName").ToString())
                If dlg.ShowDialog(Me) = DialogResult.OK Then
                    RefreshGrids()
                End If
            End Using
        End Sub

        Private Sub OnMoveStorage(sender As Object, e As EventArgs)
            If MainFrm.CurrentUser Is Nothing Then
                MessageBox.Show("Authentication Required.", "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim currentLoc = String.Format("{0}/{1}/{2}", DocRow("CabinetID"), DocRow("ShelfNo"), DocRow("BoxCode"))
            Using dlg As New FormMoveStorage(DocID, currentLoc, MainFrm.CurrentUser("FullName").ToString())
                If dlg.ShowDialog(Me) = DialogResult.OK Then
                    RefreshGrids()
                End If
            End Using
        End Sub

        Private Sub OnLaunchPDF(sender As Object, e As EventArgs)
            Dim url = DocRow("GDriveURL").ToString()
            Dim errUrl As String = ""
            If Not EmbeddedDB.ValidateGDriveURL(url, errUrl) Then
                MessageBox.Show(errUrl, "Security Validation Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            Try
                Process.Start(New ProcessStartInfo With {.FileName = url, .UseShellExecute = True})
                If MainFrm.CurrentUser IsNot Nothing Then
                    EmbeddedDB.LogAudit(MainFrm.CurrentUser("FullName").ToString(), "Launched Google Drive Soft Copy PDF: " & url)
                End If
            Catch ex As Exception
                MessageBox.Show("Error opening URL: " & ex.Message, "Launch Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub
    End Class
End Namespace
