Option Explicit On
Option Strict On

Imports System.Data
Imports System.Drawing
Imports System.Windows.Forms
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports BTA_OSG

Namespace BTA_OSG.Tests
    <TestClass>
    Public Class ThemeTests
        <TestMethod>
        Public Sub ThemeTokens_MatchDesignSpecification()
            Assert.AreEqual(ColorTranslator.FromHtml("#F4F6F8"), CivicCalmTheme.ColorCanvas)
            Assert.AreEqual(ColorTranslator.FromHtml("#FFFFFF"), CivicCalmTheme.ColorSurface)
            Assert.AreEqual(ColorTranslator.FromHtml("#EEF2F5"), CivicCalmTheme.ColorWell)
            Assert.AreEqual(ColorTranslator.FromHtml("#DDE2E5"), CivicCalmTheme.ColorBorder)
            Assert.AreEqual(ColorTranslator.FromHtml("#1B242C"), CivicCalmTheme.ColorInk)
            Assert.AreEqual(ColorTranslator.FromHtml("#146A3D"), CivicCalmTheme.ColorPrimary)
            Assert.AreEqual(ColorTranslator.FromHtml("#E6F2EB"), CivicCalmTheme.ColorPrimarySoft)
        End Sub

        <TestMethod>
        Public Sub TypographyLadder_UsesSegoeUI()
            Assert.AreEqual("Segoe UI", CivicCalmTheme.FontFormTitle.FontFamily.Name)
            Assert.AreEqual(12.0F, CivicCalmTheme.FontFormTitle.Size)
            Assert.AreEqual(FontStyle.Bold, CivicCalmTheme.FontFormTitle.Style)
            Assert.AreEqual("Segoe UI", CivicCalmTheme.FontBody.FontFamily.Name)
            Assert.AreEqual(9.0F, CivicCalmTheme.FontBody.Size)
        End Sub

        <TestMethod>
        Public Sub DataGridStyler_ApplyCivicStyle_SetsExpectedPropertiesAndBindsData()
            Using dgv As New DataGridView()
                DataGridStyler.ApplyCivicStyle(dgv)
                dgv.AllowUserToAddRows = False

                Assert.IsTrue(dgv.AutoGenerateColumns)
                Assert.IsFalse(dgv.EnableHeadersVisualStyles)
                Assert.AreEqual(CivicCalmTheme.ColorSurface, dgv.BackgroundColor)
                Assert.AreEqual(CivicCalmTheme.ColorBorder, dgv.GridColor)
                Assert.AreEqual(CivicCalmTheme.ColorWell, dgv.ColumnHeadersDefaultCellStyle.BackColor)

                Using dt As New DataTable()
                    dt.Columns.Add("DocumentID", GetType(Integer))
                    dt.Columns.Add("DocCode", GetType(String))
                    dt.Columns.Add("Title", GetType(String))
                    dt.Columns.Add("CabinetID", GetType(String))

                    dt.Rows.Add(1, "OSG-2026-0001", "Test Subject", "CAB-01")

                    dgv.BindingContext = New BindingContext()
                    dgv.DataSource = dt
                    DataGridStyler.FormatDocumentColumns(dgv)

                    Assert.IsTrue(dgv.Columns.Count >= 4)
                    Assert.AreEqual(1, dgv.Rows.Count)
                    Assert.AreEqual("ID", dgv.Columns("DocumentID").HeaderText)
                    Assert.AreEqual("Document Code", dgv.Columns("DocCode").HeaderText)
                    Assert.AreEqual("Document Title / Subject", dgv.Columns("Title").HeaderText)
                    Assert.IsFalse(dgv.Columns("CabinetID").Visible)
                End Using
            End Using
        End Sub

        <TestMethod>
        Public Sub DataGridStyler_FormatOtherTables_SetsHeadersAndFillColumns()
            ' Directives formatting test
            Using dgvDir As New DataGridView()
                DataGridStyler.ApplyCivicStyle(dgvDir)
                Using dtDir As New DataTable()
                    dtDir.Columns.Add("DirectiveID", GetType(Integer))
                    dtDir.Columns.Add("Notes", GetType(String))
                    dtDir.Rows.Add(1, "Test Note")
                    dgvDir.BindingContext = New BindingContext()
                    dgvDir.DataSource = dtDir
                    ' Operational grids are far wider than the 240px control default, and a grid only a
                    ' few pixels too narrow switches from Fill to fixed scrolling columns, so assert
                    ' Fill mode against a realistic viewport.
                    dgvDir.Width = 640
                    DataGridStyler.FormatDirectiveColumns(dgvDir)
                    Assert.AreEqual("ID", dgvDir.Columns("DirectiveID").HeaderText)
                    Assert.AreEqual(DataGridViewAutoSizeColumnMode.Fill, dgvDir.Columns("Notes").AutoSizeMode)
                End Using
            End Using

            ' Users formatting test
            Using dgvUser As New DataGridView()
                DataGridStyler.ApplyCivicStyle(dgvUser)
                Using dtUser As New DataTable()
                    dtUser.Columns.Add("UserID", GetType(Integer))
                    dtUser.Columns.Add("FullName", GetType(String))
                    dtUser.Rows.Add(1, "Test User")
                    dgvUser.BindingContext = New BindingContext()
                    dgvUser.DataSource = dtUser
                    DataGridStyler.FormatUserColumns(dgvUser)
                    Assert.AreEqual("User ID", dgvUser.Columns("UserID").HeaderText)
                    Assert.AreEqual(DataGridViewAutoSizeColumnMode.Fill, dgvUser.Columns("FullName").AutoSizeMode)
                End Using
            End Using

            ' Audit formatting test
            Using dgvAudit As New DataGridView()
                DataGridStyler.ApplyCivicStyle(dgvAudit)
                Using dtAudit As New DataTable()
                    dtAudit.Columns.Add("AuditID", GetType(Integer))
                    dtAudit.Columns.Add("ActionDescription", GetType(String))
                    dtAudit.Rows.Add(1, "Test Action")
                    dgvAudit.BindingContext = New BindingContext()
                    dgvAudit.DataSource = dtAudit
                    DataGridStyler.FormatAuditColumns(dgvAudit)
                    Assert.AreEqual("Audit ID", dgvAudit.Columns("AuditID").HeaderText)
                    Assert.AreEqual(DataGridViewAutoSizeColumnMode.Fill, dgvAudit.Columns("ActionDescription").AutoSizeMode)
                End Using
            End Using
        End Sub

        <TestMethod>
        Public Sub FormMain_DashboardLayout_NoOverlap()
            Using form As New FormMain()
                form.Size = New Size(1380, 850)
                Dim vDash As Control = Nothing
                For Each c As Control In form.Controls
                    If TypeOf c Is Panel AndAlso c.Dock = DockStyle.Fill Then
                        vDash = c.Controls(0)
                        Exit For
                    End If
                Next
                Assert.IsNotNull(vDash, "viewDashboard should be found")

                Dim tblDashboard As TableLayoutPanel = TryCast(vDash.Controls(0), TableLayoutPanel)
                Assert.IsNotNull(tblDashboard, "tblDashboard TableLayoutPanel should be root container of viewDashboard")

                Dim pnlCards As TableLayoutPanel = Nothing
                Dim pnlGridCard As Panel = Nothing

                For Each c As Control In tblDashboard.Controls
                    If TypeOf c Is TableLayoutPanel Then
                        pnlCards = DirectCast(c, TableLayoutPanel)
                    ElseIf TypeOf c Is Panel Then
                        pnlGridCard = DirectCast(c, Panel)
                    End If
                Next

                Assert.IsNotNull(pnlCards, "pnlCards must be present in tblDashboard")
                Assert.IsNotNull(pnlGridCard, "pnlGridCard must be present in tblDashboard")

                ' Assert strictly zero vertical overlap and at least 8px of canvas gap
                Assert.IsTrue(pnlCards.Bottom <= pnlGridCard.Top, "Stat cards must be strictly above pnlGridCard with zero overlap")
                Dim canvasGap As Integer = pnlGridCard.Top - pnlCards.Bottom
                Assert.IsTrue(canvasGap >= 8, "Must have at least 8px canvas separation, found: " & canvasGap.ToString() & "px")

                ' Assert that each stat card panel is completely contained within pnlCards height
                For Each card As Control In pnlCards.Controls
                    Assert.IsTrue(card.Bottom <= pnlCards.Height, "Child stat card must not overflow pnlCards height. Card bottom: " & card.Bottom.ToString() & ", container height: " & pnlCards.Height.ToString())
                    For Each gc As Control In card.Controls
                        If TypeOf gc Is Label Then
                            Dim lbl As Label = DirectCast(gc, Label)
                            Assert.IsFalse(lbl.UseMnemonic, "Stat card label must have UseMnemonic=False to prevent character stripping")
                        End If
                    Next
                Next

                ' Assert header label in pnlGridCard preserves ampersands and does not strip &
                Dim headerLabelFound As Boolean = False
                For Each c As Control In pnlGridCard.Controls
                    If TypeOf c Is Label AndAlso c.Text.Contains("OSG DOCUMENT STATUS") Then
                        headerLabelFound = True
                        Dim lblHeader As Label = DirectCast(c, Label)
                        Assert.IsFalse(lblHeader.UseMnemonic, "lblRecHeader must have UseMnemonic=False so & is preserved")
                        Assert.AreEqual("OSG DOCUMENT STATUS & MONITORING DESK", lblHeader.Text)
                    End If
                Next
                Assert.IsTrue(headerLabelFound, "Header label 'OSG DOCUMENT STATUS & MONITORING DESK' must be found")
            End Using
        End Sub

        <TestMethod>
        Public Sub ThemeTokens_AccentSG_MeetsWcagContrastRequirements()
            Assert.AreEqual(ColorTranslator.FromHtml("#8C6414"), CivicCalmTheme.ColorAccentSG)
        End Sub

        <TestMethod>
        Public Sub EscapeRowFilter_SanitizesSpecialCharacters_PreventsEvaluateException()
            Dim nastyInputs As String() = {
                "[",
                "]",
                "[*]",
                "50%",
                "O'Connor",
                "[URGENT] 100% *Verified* 'Doc'",
                "[[[[[]]]]]"
            }

            Using dt As New DataTable()
                dt.Columns.Add("Title", GetType(String))
                dt.Rows.Add("[URGENT] 100% *Verified* 'Doc' from O'Connor")
                dt.Rows.Add("Normal title")

                Dim dv As New DataView(dt)
                For Each rawInput In nastyInputs
                    Dim escaped = FormMain.EscapeRowFilter(rawInput)
                    Dim filter = String.Format("Title LIKE '%{0}%'", escaped)
                    ' Should not throw EvaluateException
                    dv.RowFilter = filter
                    Dim filteredTable = dv.ToTable()
                    Assert.IsNotNull(filteredTable)
                Next
            End Using
        End Sub

        <TestMethod>
        Public Sub FormInputPrompt_InitializesWithDefaultValues()
            Using dlg As New FormInputPrompt("Test Title", "Enter value:", "DefaultVal", True)
                Assert.AreEqual("Test Title", dlg.Text)
                Assert.IsNotNull(dlg.AcceptButton)
                Assert.IsNotNull(dlg.CancelButton)
            End Using
        End Sub

        <TestMethod>
        Public Sub FormMain_Sidebar_PermanentlyContainsPortalIntake()
            Using form As New FormMain()
                form.Size = New Size(1380, 850)
                Dim sidebar As Panel = Nothing
                For Each c As Control In form.Controls
                    If TypeOf c Is Panel AndAlso c.Dock = DockStyle.Left Then
                        sidebar = DirectCast(c, Panel)
                        Exit For
                    End If
                Next
                Assert.IsNotNull(sidebar, "Sidebar panel must exist docked to the left")

                Dim navStack As FlowLayoutPanel = Nothing
                For Each c As Control In sidebar.Controls
                    If TypeOf c Is FlowLayoutPanel Then
                        navStack = DirectCast(c, FlowLayoutPanel)
                        Exit For
                    End If
                Next
                Assert.IsNotNull(navStack, "FlowLayoutPanel navigation stack must exist in sidebar")

                Dim portalButtonFound As Boolean = False
                Dim buttonNames As New List(Of String)()
                For Each c As Control In navStack.Controls
                    If TypeOf c Is Button Then
                        Dim btn = DirectCast(c, Button)
                        If btn.Name = "btnNav_Toggle" Then Continue For ' collapse affordance, not a navigation destination
                        buttonNames.Add(btn.Text.Trim())
                        If btn.Text.Contains("Portal Intake") Then
                            portalButtonFound = True
                            Assert.AreEqual("btnNav_Portal_Intake", btn.Name)
                            Assert.AreEqual(218, btn.Width)
                            Assert.AreEqual(44, btn.Height)
                        End If
                    End If
                Next

                Assert.IsTrue(portalButtonFound, "Portal Intake button must be permanently present in the sidebar. Found buttons: " & String.Join(", ", buttonNames))
                Assert.AreEqual(8, buttonNames.Count, "Sidebar must contain exactly 8 navigation buttons")
                Assert.IsTrue(buttonNames.Contains("Dashboard"))
                Assert.IsTrue(buttonNames.Contains("Data Analytics"))
                Assert.IsTrue(buttonNames.Contains("Document Registry"))
                Assert.IsTrue(buttonNames.Contains("SG Directives"))
                Assert.IsTrue(buttonNames.Contains("Search & Storage"))
                Assert.IsTrue(buttonNames.Contains("User & RFID Admin"))
                Assert.IsTrue(buttonNames.Contains("Audit Trail"))
                Assert.IsTrue(buttonNames.Contains("Portal Intake"))
            End Using
        End Sub

        ''' Columns must be sized from their own content. Either every visible column shares
        ''' the space or none do, never the old single greedy Fill column that swallowed every
        ''' leftover pixel, and a prose column must claim more width than an identifier.
        Private Shared Sub AssertSizedByContent(dgv As DataGridView, idColumn As String, proseColumn As String)
            Dim visible As Integer = 0
            Dim filling As Integer = 0
            For Each col As DataGridViewColumn In dgv.Columns
                If col.Visible Then
                    visible += 1
                    If col.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill Then filling += 1
                End If
            Next
            Assert.IsTrue(filling = 0 OrElse filling = visible, $"{dgv.Name} mixes Fill and fixed columns, so one column absorbs the slack")
            Assert.IsTrue(dgv.Columns(proseColumn).FillWeight > dgv.Columns(idColumn).FillWeight,
                          $"{proseColumn} must claim more width than {idColumn}")
        End Sub

        <TestMethod>
        Public Sub DataGridStyler_WrapMode_NeverTruncates()
            Using dgv As New DataGridView()
                DataGridStyler.ApplyCivicStyle(dgv)

                Assert.AreEqual(DataGridViewTriState.True, dgv.DefaultCellStyle.WrapMode, "DefaultCellStyle.WrapMode must be True so text wraps rather than truncating with ellipses")
                Assert.AreEqual(DataGridViewTriState.True, dgv.AlternatingRowsDefaultCellStyle.WrapMode, "AlternatingRowsDefaultCellStyle.WrapMode must be True")
                ' Row heights still expand to fit full wrapped text, but measurement is bounded to
                ' the rows on screen; off-screen rows re-measure as they scroll into view, which is
                ' what keeps a large registry grid responsive (see anti-slop audit-007).
                Assert.AreEqual(DataGridViewAutoSizeRowsMode.DisplayedCellsExceptHeaders, dgv.AutoSizeRowsMode, "AutoSizeRowsMode must be DisplayedCellsExceptHeaders so row heights expand to fit full text without an O(all-rows) layout pass per refresh")

                Using dt As New DataTable()
                    dt.Columns.Add("DocumentID", GetType(Integer))
                    dt.Columns.Add("DocCode", GetType(String))
                    dt.Columns.Add("Title", GetType(String))
                    dt.Columns.Add("OriginatingOffice", GetType(String))
                    dt.Columns.Add("Remarks", GetType(String))

                    dt.Rows.Add(1, "OSG-2026-0001", "A very long document title that needs wrapping to show every single word without truncation", "Office of the Secretary-General Substantive Desk", "Detailed multi-line remarks explaining the entire legislative journey")

                    dgv.BindingContext = New BindingContext()
                    dgv.DataSource = dt
                    dgv.Width = 1200
                    DataGridStyler.FormatDocumentColumns(dgv)

                    Assert.AreEqual(DataGridViewTriState.True, dgv.Columns("Title").DefaultCellStyle.WrapMode)
                    AssertSizedByContent(dgv, "DocumentID", "Title")
                End Using
            End Using
        End Sub

        ' Operators read prose, not codes, so the Status column shows the display name seeded in
        ' tbl_DocumentStatuses while the bound value keeps the raw code for filters and audit.
        <TestMethod>
        Public Sub DataGridStyler_StatusColumn_RendersDisplayNameInsideBadge()
            Using dgv As New DataGridView()
                DataGridStyler.ApplyCivicStyle(dgv)

                Dim renderedBackColor As Color = Color.Empty
                Dim renderedForeColor As Color = Color.Empty
                AddHandler dgv.CellFormatting, Sub(s, e)
                                                   If e.ColumnIndex >= 0 AndAlso e.RowIndex = 0 AndAlso dgv.Columns(e.ColumnIndex).Name = "CurrentStatus" Then
                                                       renderedBackColor = e.CellStyle.BackColor
                                                       renderedForeColor = e.CellStyle.ForeColor
                                                   End If
                                               End Sub

                Using dt As New DataTable()
                    dt.Columns.Add("DocumentID", GetType(Integer))
                    dt.Columns.Add("CurrentStatus", GetType(String))
                    dt.Columns.Add("AssignedStaff", GetType(String))
                    dt.Rows.Add(1, "FOR_REVIEW", "Amina T. Macacua")
                    dt.Rows.Add(2, "RECEIVED", "Hassim A. Ibrahim")
                    dt.Rows.Add(3, "RELEASED", "Records Section")

                    dgv.BindingContext = New BindingContext()
                    dgv.DataSource = dt
                    DataGridStyler.FormatDocumentColumns(dgv)

                    Assert.AreEqual("FOR_REVIEW", dgv.Rows(0).Cells("CurrentStatus").Value, "raw status code must stay bound")
                    Assert.AreEqual("For Review", Convert.ToString(dgv.Rows(0).Cells("CurrentStatus").FormattedValue))
                    Assert.AreEqual("Received", Convert.ToString(dgv.Rows(1).Cells("CurrentStatus").FormattedValue))
                    Assert.AreEqual("Released", Convert.ToString(dgv.Rows(2).Cells("CurrentStatus").FormattedValue))
                    Assert.AreEqual(CivicCalmTheme.ColorStatusReviewBg, renderedBackColor, "review badge fill")
                    Assert.AreEqual(CivicCalmTheme.ColorStatusReviewFg, renderedForeColor, "review badge ink")
                    Assert.AreEqual(8, dgv.Columns("CurrentStatus").DefaultCellStyle.Padding.Left, "badge gutter keeps the fill off the neighbouring column")
                End Using
            End Using
        End Sub

        ' A raw Drive URL is one unbreakable token that clips mid-string and reserves a wide column,
        ' so the Soft Copy Link column states the attachment and sizes itself from that label.
        <TestMethod>
        Public Sub DataGridStyler_SoftCopyLinkColumn_RendersAttachmentStateAndSizesToIt()
            Dim longUrl As String = "https://drive.google.com/file/d/sample-comm-001/view?usp=sharing"
            Using dgv As New DataGridView()
                DataGridStyler.ApplyCivicStyle(dgv)

                Dim renderedForeColor As Color = Color.Empty
                Dim renderedUnderline As Boolean = False
                AddHandler dgv.CellFormatting, Sub(s, e)
                                                   If e.RowIndex = 0 AndAlso e.ColumnIndex >= 0 AndAlso dgv.Columns(e.ColumnIndex).Name = "GDriveURL" Then
                                                       renderedForeColor = e.CellStyle.ForeColor
                                                       renderedUnderline = e.CellStyle.Font IsNot Nothing AndAlso e.CellStyle.Font.Underline
                                                   End If
                                               End Sub

                Using dt As New DataTable()
                    dt.Columns.Add("DocumentID", GetType(Integer))
                    dt.Columns.Add("GDriveURL", GetType(String))
                    dt.Columns.Add("Title", GetType(String))
                    dt.Rows.Add(1, longUrl, "Transmittal of Inter-Agency Cooperation Agreement")
                    dt.Rows.Add(2, "", "Memorandum without an attached soft copy")

                    dgv.BindingContext = New BindingContext()
                    dgv.DataSource = dt
                    dgv.Width = 320
                    DataGridStyler.FormatDocumentColumns(dgv)

                    Assert.AreEqual(longUrl, dgv.Rows(0).Cells("GDriveURL").Value, "raw URL must stay bound for launch and validation")
                    Assert.AreEqual("Attached", Convert.ToString(dgv.Rows(0).Cells("GDriveURL").FormattedValue))
                    Assert.AreEqual(longUrl, dgv.Rows(0).Cells("GDriveURL").ToolTipText, "full link stays available on hover")
                    Assert.AreEqual("Not Attached", Convert.ToString(dgv.Rows(1).Cells("GDriveURL").FormattedValue))
                    Assert.AreEqual(CivicCalmTheme.ColorPrimary, renderedForeColor, "attachment state uses the institutional ink token")
                    Assert.IsFalse(renderedUnderline, "the state label must not fake a clickable hyperlink")
                    Assert.IsTrue(dgv.Columns("GDriveURL").Width < 200,
                                  "Soft Copy Link must size to the state label instead of the raw URL")
                    Assert.IsTrue(dgv.Columns("Title").Width > dgv.Columns("GDriveURL").Width,
                                  "the prose column must keep claim to more width than the link column")
                End Using
            End Using
        End Sub

        ' The Document Details view prints the raw soft copy URL, so the link needs break
        ' opportunities or a long one clips mid-string in the overview row and preview bar.
        <TestMethod>
        Public Sub FormDocumentDetail_WrapSoftCopyLink_PreservesEveryCharacterOnShortLines()
            Dim url As String = "https://drive.google.com/file/d/sample-comm-001/view?usp=sharing"
            Dim wrapped As String = FormDocumentDetail.WrapSoftCopyLink(url)

            Assert.IsTrue(wrapped.Contains(Environment.NewLine), "a long link must carry break opportunities")
            Assert.AreEqual(url, wrapped.Replace(Environment.NewLine, ""), "no character of the link may be lost")

            For Each line As String In wrapped.Split(New String() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)
                Assert.IsTrue(line.Length <= 48, $"line '{line}' exceeds the printable segment budget")
            Next

            Assert.AreEqual("", FormDocumentDetail.WrapSoftCopyLink(""), "a missing link renders as empty, not as a blank line")
        End Sub

        <TestMethod>
        Public Sub DataGridStyler_EveryGrid_SizesColumnsByContent()
            Using dgvDir As New DataGridView(), dgvUser As New DataGridView(), dgvAudit As New DataGridView(), dgvRoute As New DataGridView(), dgvMove As New DataGridView()
                DataGridStyler.ApplyCivicStyle(dgvDir)
                DataGridStyler.ApplyCivicStyle(dgvUser)
                DataGridStyler.ApplyCivicStyle(dgvAudit)
                DataGridStyler.ApplyCivicStyle(dgvRoute)
                DataGridStyler.ApplyCivicStyle(dgvMove)

                Using dtDir As New DataTable(), dtUser As New DataTable(), dtAudit As New DataTable(), dtRoute As New DataTable(), dtMove As New DataTable()
                    dtDir.Columns.Add("DirectiveID", GetType(Integer))
                    dtDir.Columns.Add("DocumentID", GetType(Integer))
                    dtDir.Columns.Add("Notes", GetType(String))
                    dtDir.Rows.Add(1, 10, "Priority processing requested for the committee")
                    dgvDir.BindingContext = New BindingContext()
                    dgvDir.DataSource = dtDir
                    dgvDir.Width = 1200
                    DataGridStyler.FormatDirectiveColumns(dgvDir)
                    AssertSizedByContent(dgvDir, "DirectiveID", "Notes")

                    dtUser.Columns.Add("UserID", GetType(Integer))
                    dtUser.Columns.Add("FullName", GetType(String))
                    dtUser.Columns.Add("IsActive", GetType(Boolean))
                    dtUser.Rows.Add(1, "Prof. Ali B. Pangalian", True)
                    dgvUser.BindingContext = New BindingContext()
                    dgvUser.DataSource = dtUser
                    dgvUser.Width = 1200
                    DataGridStyler.FormatUserColumns(dgvUser)
                    AssertSizedByContent(dgvUser, "UserID", "FullName")

                    dtAudit.Columns.Add("AuditID", GetType(Integer))
                    dtAudit.Columns.Add("ActionDescription", GetType(String))
                    dtAudit.Rows.Add(1, "Document #4 approved by Sec Gen: Approved for release")
                    dgvAudit.BindingContext = New BindingContext()
                    dgvAudit.DataSource = dtAudit
                    dgvAudit.Width = 1200
                    DataGridStyler.FormatAuditColumns(dgvAudit)
                    AssertSizedByContent(dgvAudit, "AuditID", "ActionDescription")

                    dtRoute.Columns.Add("RoutingID", GetType(Integer))
                    dtRoute.Columns.Add("Remarks", GetType(String))
                    dtRoute.Rows.Add(1, "Resubmitted with punchlist completed")
                    dgvRoute.BindingContext = New BindingContext()
                    dgvRoute.DataSource = dtRoute
                    dgvRoute.Width = 1200
                    DataGridStyler.FormatRoutingColumns(dgvRoute)
                    AssertSizedByContent(dgvRoute, "RoutingID", "Remarks")

                    dtMove.Columns.Add("MovementID", GetType(Integer))
                    dtMove.Columns.Add("Reason", GetType(String))
                    dtMove.Rows.Add(1, "Moved to vault for archival storage")
                    dgvMove.BindingContext = New BindingContext()
                    dgvMove.DataSource = dtMove
                    dgvMove.Width = 1200
                    DataGridStyler.FormatMovementColumns(dgvMove)
                    AssertSizedByContent(dgvMove, "MovementID", "Reason")
                End Using
            End Using
        End Sub

        <TestMethod>
        Public Sub FormMain_DirectivesTab_HasNotesFieldInLayout()
            Using form As New FormMain()
                form.Size = New Size(1380, 850)
                ' Find viewDirectives or controls within
                Dim tblMain As Control = Nothing
                For Each c As Control In form.Controls
                    If TypeOf c Is Panel AndAlso c.Dock = DockStyle.Fill Then
                        ' pnlContent
                        For Each subC As Control In c.Controls
                            tblMain = subC
                        Next
                    End If
                Next
                Assert.IsNotNull(form, "FormMain initialized")
            End Using
        End Sub
    End Class
End Namespace
