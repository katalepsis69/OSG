Option Explicit On
Option Strict On

Imports System
Imports System.Data
Imports System.Drawing
Imports System.Windows.Forms

Namespace BTA_OSG
    Partial Public Class FormDocumentDetail
        Inherits Form

        Private DocID As Integer
        Private MainFrm As FormMain
        Private DocRow As DataRow

        Private pnlHeader As Panel
        Private lblCode As Label
        Private lblTitle As Label
        Private lblStatusBadge As Label

        Private btnRoutingSlip As Button
        Private pnlWorkflowRibbon As Panel
        Private lblStationBadge As Label
        Private lblNextStationBadge As Label
        Private lblProgressBadge As Label
        Private dgvSteps As DataGridView

        Private pnlPunchlistBanner As Panel
        Private lblPunchlistContent As Label

        Private btnRequestRevision As Button
        Private btnResubmit As Button
        Private btnApprove As Button
        Private btnRelease As Button
        Private btnRoute As Button
        Private btnMove As Button
        Private btnLaunchPdf As Button
        Private btnClose As Button

        Private tabDetail As TabControl
        Private tabOverview As TabPage
        Private tabDirectives As TabPage
        Private tabRouting As TabPage
        Private tabMovements As TabPage
        Private tabPreview As TabPage

        Private dgvDirectives As DataGridView
        Private dgvRouting As DataGridView
        Private dgvMovements As DataGridView
        Private wvPreview As Microsoft.Web.WebView2.WinForms.WebView2
        Private wvEnvironment As Microsoft.Web.WebView2.Core.CoreWebView2Environment
        Private lblPreviewPlaceholder As Label
        Private lblDirectivesWatermark As Label
        Private lblRoutingWatermark As Label
        Private lblMovementsWatermark As Label
        Private lblStepsWatermark As Label

        Private pnlOverviewTable As TableLayoutPanel

        ' Tracks the URL the embedded preview last navigated to so repeat visits to the preview
        ' tab do not reload the page.
        Private previewNavigatedUrl As String = Nothing

        ' This dialog deliberately does NOT use UiBuffering.WsExComposited, unlike the main
        ' window. Compositing renders the whole child tree through one redirection surface, and
        ' the embedded WebView2 preview paints its own Chromium surface into that same window:
        ' while the preview is on screen the header and ribbon siblings re-present on every web
        ' frame, which reads as blinking. Panel-level double buffering (EnableDeep below) stays
        ' on, so ordinary repaint flicker is still covered. The cost is that maximize and
        ' restore can ghost the relaid-out action buttons, which this dialog rarely sees.
        Public Sub New(id As Integer, parentForm As FormMain)
            DocID = id
            MainFrm = parentForm
            InitializeForm()
        End Sub

        Private Sub InitializeForm()
            If Not LoadDocData() Then Return
            AppAssets.ApplyFormIcon(Me)

            Me.Text = String.Format("Document Details & Specifications : [{0}] {1}", DocRow("DocCode"), DocRow("Title"))
            Me.Size = New Size(1180, 820)
            Me.MinimumSize = New Size(980, 680)
            Me.StartPosition = FormStartPosition.CenterParent
            Me.Font = CivicCalmTheme.FontBody
            Me.BackColor = CivicCalmTheme.ColorCanvas

            SetupHeaderPanel()
            SetupWorkflowRibbon()
            SetupPunchlistBanner()
            SetupTabControl()
            AddHandler tabDetail.Selected, Sub(s As Object, e As TabControlEventArgs)
                                               If e.TabPage Is tabPreview Then LoadPreviewTab()
                                           End Sub

            Me.Controls.Add(tabDetail)
            Me.Controls.Add(pnlPunchlistBanner)
            Me.Controls.Add(pnlWorkflowRibbon)
            Me.Controls.Add(pnlHeader)

            ' Buffer every container so maximize/restore repaints atomically (same treatment as
            ' the main window).
            UiBuffering.EnableDeep(Me)

            RefreshGrids()

            ' RefreshGrids balanced the grids before this form was laid out, so first open shows
            ' widths measured against pre-layout bounds; re-balance once the client size is real.
            ' dgvSteps is included: it is filled in PopulateStepsGrid during the pre-layout pass
            ' and needs the same post-layout measurement (first-click render bug class).
            AddHandler Me.Shown, Sub()
                                     DataGridStyler.FormatDirectiveColumns(dgvDirectives)
                                     DataGridStyler.FormatRoutingColumns(dgvRouting)
                                     DataGridStyler.FormatMovementColumns(dgvMovements)
                                     dgvSteps.AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCells)
                                 End Sub
        End Sub

        Private Function LoadDocData() As Boolean
            Dim rows As DataRow()
            ' The replay writer mutates these tables on the sync thread; read under the
            ' store lock like every other cross-thread select.
            SyncLock EmbeddedDB.SyncRoot
                rows = EmbeddedDB.DataSet.Tables("Documents").Select("DocumentID = " & DocID)
            End SyncLock
            If rows.Length = 0 Then
                MessageBox.Show("Document record not found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.Close()
                Return False
            End If
            DocRow = rows(0)
            Return True
        End Function

        Private Sub RefreshGrids()
            If Not LoadDocData() Then Return

            lblCode.Text = String.Format("DOC CODE: {0}   |   TYPE: {1}   |   FLOW: {2}", DocRow("DocCode"), DocRow("DocType").ToString().ToUpperInvariant(), DocRow("FlowDirection"))
            lblTitle.Text = DocRow("Title").ToString()
            lblStatusBadge.Text = String.Format("Status: {0}   |   Section: {1}   |   Staff: {2}", DocumentStatus.DisplayName(DocRow("CurrentStatus").ToString()), DocRow("AssignedSection"), DocRow("AssignedStaff"))

            Dim punchlist = DocRow("RevisionPunchlist").ToString().Trim()
            If Not String.IsNullOrEmpty(punchlist) Then
                lblPunchlistContent.Text = punchlist
                ' Size the banner to the wrapped text so a long punchlist is not clipped by
                ' the original fixed 70px height; 44 covers the title row, padding, border.
                pnlPunchlistBanner.Height = Math.Max(70,
                    TextRenderer.MeasureText(punchlist, CivicCalmTheme.FontBody,
                        New Size(Math.Max(400, Me.ClientSize.Width - 72), 0),
                        TextFormatFlags.WordBreak).Height + 44)
                pnlPunchlistBanner.Visible = True
            Else
                pnlPunchlistBanner.Visible = False
            End If

            UpdateActionButtons()
            SetupOverviewTab()

            ' Preview navigation is deferred until the operator actually opens the preview tab;
            ' a refresh after an action only re-renders it when that tab is already on screen.
            If tabDetail.SelectedTab Is tabPreview Then LoadPreviewTab()

            Dim dvDir As New DataView(EmbeddedDB.DataSet.Tables("Directives"))
            dvDir.RowFilter = "DocumentID = " & DocID
            dgvDirectives.DataSource = dvDir
            DataGridStyler.FormatDirectiveColumns(dgvDirectives)
            ShowGridState(dgvDirectives, lblDirectivesWatermark, "No SG directives have been issued for this document.")

            Dim dvRoute As New DataView(EmbeddedDB.DataSet.Tables("RoutingLogs"))
            dvRoute.RowFilter = "DocumentID = " & DocID
            dgvRouting.DataSource = dvRoute
            DataGridStyler.FormatRoutingColumns(dgvRouting)
            ShowGridState(dgvRouting, lblRoutingWatermark, "No transmittal logs yet. Route the document to begin its custody trail.")

            Dim dvMove As New DataView(EmbeddedDB.DataSet.Tables("Movements"))
            dvMove.RowFilter = "DocumentID = " & DocID
            dgvMovements.DataSource = dvMove
            DataGridStyler.FormatMovementColumns(dgvMovements)
            ShowGridState(dgvMovements, lblMovementsWatermark, "The storage landmark has not moved since registration.")

            ' Calculate Step-by-Step Pipeline
            Dim currentDoc = GetCurrentDocumentObject()
            Dim currentLogs = GetCurrentRoutingLogsList()
            Dim steps = RoutingStepService.GetStepsForDocument(currentDoc, currentLogs)

            Dim curStation = RoutingStepService.GetCurrentStation(steps)
            Dim nxtStation = RoutingStepService.GetNextStation(steps)
            Dim pct = RoutingStepService.GetProgressPercentage(steps)

            If curStation IsNot Nothing Then
                lblStationBadge.Text = String.Format("CURRENT STATION: {0} ({1})", curStation.ResponsibleOffice, curStation.StageName)
            Else
                lblStationBadge.Text = "CURRENT STATION: Completed"
            End If

            If nxtStation IsNot Nothing Then
                lblNextStationBadge.Text = String.Format("Next Step: {0} ({1})", nxtStation.StageName, nxtStation.ResponsibleOffice)
            Else
                lblNextStationBadge.Text = "Next Step: Completed / Archived"
            End If

            lblProgressBadge.Text = String.Format("{0}% Complete", pct)
            PopulateStepsGrid(steps)
        End Sub

        ' The embedded browser pays its navigation cost only when the operator actually opens
        ' the preview tab; re-entry with the same URL keeps the page that is already rendered.
        ' WebView2 (Edge) replaces the retired IE-based WebBrowser control, whose engine
        ' cannot run Google Drive's preview.
        Private Async Sub LoadPreviewTab()
            ' Deny-by-default: a dropped session or a NULL flag must not expose the preview.
            Dim canSoftCopyUser As Boolean = False
            If MainFrm.CurrentUser IsNot Nothing AndAlso MainFrm.CurrentUser.Table.Columns.Contains("CanSoftCopy") AndAlso Not IsDBNull(MainFrm.CurrentUser("CanSoftCopy")) Then
                canSoftCopyUser = Convert.ToBoolean(MainFrm.CurrentUser("CanSoftCopy"))
            End If

            If Not canSoftCopyUser Then
                lblPreviewPlaceholder.Text = "Access Denied: Account policy does not permit soft-copy document preview."
                lblPreviewPlaceholder.Visible = True
                wvPreview.Visible = False
                Return
            End If

            Dim url = DocRow("GDriveURL").ToString().Trim()
            ' The preview is a soft-copy surface like Launch, so it validates against the
            ' same allowlist: a mirrored row or an "All Files" attachment must not render
            ' here what Launch would refuse.
            Dim validationError As String = ""
            If String.IsNullOrEmpty(url) OrElse Not EmbeddedDB.ValidateGDriveURL(url, validationError) Then
                wvPreview.Visible = False
                lblPreviewPlaceholder.Visible = True
                lblPreviewPlaceholder.Text = If(String.IsNullOrEmpty(url),
                    "No soft-copy attachment is linked to this document.",
                    "Soft-copy preview refused: " & validationError)
                previewNavigatedUrl = Nothing
                Return
            End If

            Dim navUrl = url
            If navUrl.Contains("drive.google.com/file/d/") Then
                If navUrl.EndsWith("/view", StringComparison.OrdinalIgnoreCase) Then
                    navUrl = navUrl.Substring(0, navUrl.Length - 5) & "/preview"
                ElseIf navUrl.Contains("/view?") Then
                    navUrl = navUrl.Replace("/view?", "/preview?")
                End If
            End If

            If String.Equals(previewNavigatedUrl, navUrl, StringComparison.OrdinalIgnoreCase) Then
                lblPreviewPlaceholder.Visible = False
                wvPreview.Visible = True
                Return
            End If

            lblPreviewPlaceholder.Visible = False
            wvPreview.Visible = True
            Try
                If wvEnvironment Is Nothing Then
                    ' The profile cache lives in AppData because a Program Files install
                    ' folder is read-only and WebView2 refuses the executable directory.
                    Dim profileDir = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "BTA_OSG_DocumentTracking", "WebView2")
                    wvEnvironment = Await Microsoft.Web.WebView2.Core.CoreWebView2Environment.CreateAsync(Nothing, profileDir)
                End If
                Await wvPreview.EnsureCoreWebView2Async(wvEnvironment)
                wvPreview.CoreWebView2.Navigate(New Uri(navUrl).AbsoluteUri)
                previewNavigatedUrl = navUrl
            Catch ex As Exception
                ' Missing Evergreen Runtime (possible on Server 2022) or a failed browser
                ' process must not look like a frozen tab: name the cause and the escape hatch.
                wvPreview.Visible = False
                lblPreviewPlaceholder.Text = "Embedded preview is unavailable: " & ex.Message & " Install the Microsoft WebView2 Runtime, or use Open in External Browser."
                lblPreviewPlaceholder.Visible = True
                previewNavigatedUrl = Nothing
            End Try
        End Sub
    End Class
End Namespace
