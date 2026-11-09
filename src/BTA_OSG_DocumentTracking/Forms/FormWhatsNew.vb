Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Globalization
Imports System.Linq
Imports System.Reflection
Imports System.Windows.Forms

Namespace BTA_OSG
    Public Class FormWhatsNew
        Inherits Form

        Private Class LocalHistoryEntry
            Public Property VersionText As String = ""
            Public Property HeaderText As String = ""
            Public Property ReleaseDate As DateTime = DateTime.MinValue
            Public Property Bullets As New List(Of String)()
        End Class

        Private lblStatusBadge As Label
        Private pnlGroups As FlowLayoutPanel
        Private pnlReading As Panel
        Private btnUpdateNow As Button
        Private flpFooter As FlowLayoutPanel
        Private prgDownload As ProgressBar
        Private btnCancelDownload As Button
        Private _downloadCts As System.Threading.CancellationTokenSource

        Private _hasUpdate As Boolean = False
        Private _latestVersion As String = ""
        Private _downloadUrl As String = ""
        Private _remoteNotes As String = ""
        Private _latestPublishedAt As DateTime = DateTime.MinValue
        Private _alreadyChecked As Boolean = False
        Private _checkSucceeded As Boolean = True

        Private ReadOnly _groupViews As New List(Of MonthGroupView)()

        ' Long-form reading ladder for the changelog surface.
        Private Shared ReadOnly FontChangelogBody As New Font("Segoe UI", 9.5F, FontStyle.Regular)
        Private Shared ReadOnly FontChangelogSubheader As New Font("Segoe UI", 9.5F, FontStyle.Bold)
        Private Shared ReadOnly FontChangelogHeader As New Font("Segoe UI", 10.5F, FontStyle.Bold)
        Private Shared ReadOnly FontChangelogSpacer As New Font("Segoe UI", 4.5F, FontStyle.Regular)
        Private Shared ReadOnly FontChangelogSectionGap As New Font("Segoe UI", 8.0F, FontStyle.Regular)

        Private Const GroupHeaderHeight As Integer = 32
        Private Const GroupBodyTopOffset As Integer = 8
        Private Const GroupBodyLeftOffset As Integer = 10

        Private Class MonthGroup
            Public Property Title As String = ""
            Public Property Expanded As Boolean = False
            Public ReadOnly Sections As New List(Of ChangelogSection)()
        End Class

        Private Class MonthGroupView
            Public Group As MonthGroup
            Public Panel As Panel
            Public Header As LinkLabel
            Public Body As RichTextBox
        End Class

        Public Sub New(Optional hasUpdate As Boolean = False, Optional latestVersion As String = "", Optional downloadUrl As String = "", Optional releaseNotes As String = "", Optional latestPublishedAt As DateTime = Nothing, Optional alreadyChecked As Boolean = False, Optional checkSucceeded As Boolean = True)
            _hasUpdate = hasUpdate
            _latestVersion = latestVersion
            _downloadUrl = downloadUrl
            _remoteNotes = releaseNotes
            _latestPublishedAt = latestPublishedAt
            _alreadyChecked = alreadyChecked
            _checkSucceeded = checkSucceeded

            InitializeComponent()
            PopulateChangelog(_remoteNotes, Nothing)
        End Sub

        Private Sub InitializeComponent()
            Me.Text = "What's New - BTA OSG Document Tracking"
            Me.Size = New Size(700, 620)
            Me.MinimumSize = New Size(560, 480)
            Me.StartPosition = FormStartPosition.CenterParent
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.ShowInTaskbar = False
            Me.BackColor = CivicCalmTheme.ColorCanvas
            Me.Font = CivicCalmTheme.FontBody

            ' Padding lives on headerLayout so pnlBorderBottom spans edge-to-edge.
            Dim pnlHeader As New Panel With {
                .Dock = DockStyle.Top,
                .Height = 84,
                .BackColor = CivicCalmTheme.ColorSurface
            }

            Dim pnlBorderBottom As New Panel With {
                .Dock = DockStyle.Bottom,
                .Height = 1,
                .BackColor = CivicCalmTheme.ColorBorder
            }
            pnlHeader.Controls.Add(pnlBorderBottom)

            Dim headerLayout As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 2,
                .RowCount = 2,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(20, 14, 20, 10)
            }
            headerLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            headerLayout.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            headerLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            headerLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))

            Dim lblHeaderTitle As New Label With {
                .Text = "What's New in BTA OSG",
                .Font = CivicCalmTheme.FontFormTitle,
                .ForeColor = CivicCalmTheme.ColorInk,
                .AutoSize = True,
                .Margin = New Padding(0, 0, 0, 2)
            }
            headerLayout.Controls.Add(lblHeaderTitle, 0, 0)

            Dim currentVer = Assembly.GetExecutingAssembly().GetName().Version
            Dim verStr = If(currentVer IsNot Nothing, $"v{AppUpdateService.FormatVersion(currentVer)}", "v?")

            Dim lblHeaderSubtitle As New Label With {
                .Text = $"Current Version: {verStr}  -  Parliamentary Document Tracking System",
                .Font = CivicCalmTheme.FontMicrocopy,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .AutoSize = True,
                .Margin = New Padding(0)
            }
            headerLayout.Controls.Add(lblHeaderSubtitle, 0, 1)

            lblStatusBadge = New Label With {
                .AutoSize = False,
                .Size = New Size(190, 28),
                .TextAlign = ContentAlignment.MiddleCenter,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Anchor = AnchorStyles.Right,
                .Margin = New Padding(8, 0, 0, 0)
            }
            headerLayout.Controls.Add(lblStatusBadge, 1, 0)
            headerLayout.SetRowSpan(lblStatusBadge, 2)
            pnlHeader.Controls.Add(headerLayout)
            headerLayout.BringToFront()

            If _hasUpdate Then
                lblStatusBadge.Text = $"Update Available ({AppUpdateService.FormatTagLabel(_latestVersion)})"
                lblStatusBadge.BackColor = CivicCalmTheme.ColorStatusReceivedBg
                lblStatusBadge.ForeColor = CivicCalmTheme.ColorStatusReceivedFg
            ElseIf _alreadyChecked AndAlso Not _checkSucceeded Then
                lblStatusBadge.Text = "Update Check Unavailable"
                lblStatusBadge.BackColor = CivicCalmTheme.ColorStatusReceivedBg
                lblStatusBadge.ForeColor = CivicCalmTheme.ColorWarning
            ElseIf Not String.IsNullOrEmpty(_latestVersion) OrElse _alreadyChecked Then
                lblStatusBadge.Text = $"System Up to Date ({verStr})"
                lblStatusBadge.BackColor = CivicCalmTheme.ColorPrimarySoft
                lblStatusBadge.ForeColor = CivicCalmTheme.ColorPrimary
            Else
                lblStatusBadge.Text = "Checking updates..."
                lblStatusBadge.BackColor = CivicCalmTheme.ColorWell
                lblStatusBadge.ForeColor = CivicCalmTheme.ColorInkMuted
            End If

            ' Padding lives on flpFooter so pnlBorderTop spans edge-to-edge above the buttons.
            Dim pnlFooter As New Panel With {
                .Dock = DockStyle.Bottom,
                .Height = 60,
                .BackColor = CivicCalmTheme.ColorSurface
            }
            Dim pnlBorderTop As New Panel With {
                .Dock = DockStyle.Top,
                .Height = 1,
                .BackColor = CivicCalmTheme.ColorBorder
            }
            pnlFooter.Controls.Add(pnlBorderTop)

            flpFooter = New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.RightToLeft,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(16, 12, 16, 12)
            }

            Dim btnClose = CivicCalmTheme.CreateIconButton("&Close", margin:=New Padding(6, 0, 0, 0))
            btnClose.MinimumSize = New Size(95, 34)
            AddHandler btnClose.Click, Sub() Me.Close()
            flpFooter.Controls.Add(btnClose)

            ' Enter and Escape close the dialog; closing mid-download cancels the transfer.
            Me.AcceptButton = btnClose
            Me.CancelButton = btnClose
            AddHandler Me.FormClosing, Sub(s, e)
                                           If _downloadCts IsNot Nothing Then
                                               _downloadCts.Cancel()
                                           End If
                                       End Sub

            If _hasUpdate AndAlso Not String.IsNullOrEmpty(_downloadUrl) Then
                ShowUpdateNowButton()
            End If

            pnlFooter.Controls.Add(flpFooter)
            flpFooter.BringToFront()

            Dim pnlContent As New Panel With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(20, 14, 20, 14),
                .BackColor = CivicCalmTheme.ColorCanvas
            }

            ' 1px border frame around the reading surface; padding lives on pnlGroups so the scrollbar sits flush on the right edge.
            pnlReading = New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = CivicCalmTheme.ColorBorder,
                .Padding = New Padding(1)
            }

            pnlGroups = New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.TopDown,
                .WrapContents = False,
                .AutoScroll = True,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(16, 14, 16, 10)
            }

            pnlReading.Controls.Add(pnlGroups)
            pnlContent.Controls.Add(pnlReading)

            Me.Controls.Add(pnlContent)
            Me.Controls.Add(pnlFooter)
            Me.Controls.Add(pnlHeader)
            pnlContent.BringToFront()
            Me.PerformLayout()
        End Sub

        Protected Overrides Async Sub OnShown(e As EventArgs)
            MyBase.OnShown(e)

            If Not _alreadyChecked AndAlso String.IsNullOrEmpty(_latestVersion) Then
                Dim info = Await AppUpdateService.CheckUpdateInfoAsync()
                If Me.IsDisposed Then Return

                Dim currentVer = Assembly.GetExecutingAssembly().GetName().Version
                Dim verStr = If(currentVer IsNot Nothing, $"v{AppUpdateService.FormatVersion(currentVer)}", "v?")

                If info.Success Then
                    _latestVersion = info.LatestVersion
                    _downloadUrl = info.DownloadUrl
                    _remoteNotes = info.ReleaseNotes
                    _latestPublishedAt = info.PublishedAt
                    _hasUpdate = info.HasUpdate

                    If _hasUpdate Then
                        lblStatusBadge.Text = $"Update Available ({AppUpdateService.FormatTagLabel(_latestVersion)})"
                        lblStatusBadge.BackColor = CivicCalmTheme.ColorStatusReceivedBg
                        lblStatusBadge.ForeColor = CivicCalmTheme.ColorStatusReceivedFg
                        ShowUpdateNowButton()
                    Else
                        lblStatusBadge.Text = $"System Up to Date ({verStr})"
                        lblStatusBadge.BackColor = CivicCalmTheme.ColorPrimarySoft
                        lblStatusBadge.ForeColor = CivicCalmTheme.ColorPrimary
                    End If
                Else
                    lblStatusBadge.Text = $"System Ready ({verStr})"
                    lblStatusBadge.BackColor = CivicCalmTheme.ColorPrimarySoft
                    lblStatusBadge.ForeColor = CivicCalmTheme.ColorPrimary
                End If
            End If

            Dim pastReleases = Await AppUpdateService.GetRecentReleasesAsync()
            If Me.IsDisposed Then Return
            PopulateChangelog(_remoteNotes, pastReleases)
        End Sub

        Private Sub ShowUpdateNowButton()
            If btnUpdateNow IsNot Nothing OrElse String.IsNullOrEmpty(_downloadUrl) Then Return

            If pnlReading IsNot Nothing Then
                prgDownload = New ProgressBar With {
                    .Dock = DockStyle.Bottom,
                    .Height = 6,
                    .Visible = False
                }
                pnlReading.Controls.Add(prgDownload)
                pnlGroups.BringToFront()
            End If

            btnUpdateNow = CivicCalmTheme.CreateIconButton("&Update && Restart Now", "arrow-clockwise", backColor:=CivicCalmTheme.ColorPrimary, margin:=New Padding(6, 0, 0, 0))
            AddHandler btnUpdateNow.Click, AddressOf OnUpdateNowClicked
            flpFooter.Controls.Add(btnUpdateNow)

            btnCancelDownload = CivicCalmTheme.CreateIconButton("&Stop Download", "x", 14, minHeight:=30, margin:=New Padding(6, 0, 0, 0), visible:=False)
            AddHandler btnCancelDownload.Click, Sub()
                                                    If _downloadCts IsNot Nothing Then _downloadCts.Cancel()
                                                End Sub
            flpFooter.Controls.Add(btnCancelDownload)
        End Sub

        Private Async Sub OnUpdateNowClicked(sender As Object, e As EventArgs)
            btnUpdateNow.Enabled = False
            If prgDownload IsNot Nothing Then prgDownload.Visible = True
            _downloadCts = New System.Threading.CancellationTokenSource()
            If btnCancelDownload IsNot Nothing Then btnCancelDownload.Visible = True
            Dim reporter As New Progress(Of DownloadProgressInfo)(Sub(p)
                                                                      If Me.IsDisposed OrElse prgDownload Is Nothing Then Return
                                                                      If p.Percent >= 0 Then
                                                                          prgDownload.Style = ProgressBarStyle.Blocks
                                                                          prgDownload.Value = Math.Min(100, p.Percent)
                                                                      Else
                                                                          prgDownload.Style = ProgressBarStyle.Marquee
                                                                      End If
                                                                      If Not String.IsNullOrEmpty(p.Label) Then btnUpdateNow.Text = p.Label
                                                                  End Sub)
            If Await AppUpdateService.DownloadAndApplyAsync(_downloadUrl, Me, reporter, _downloadCts.Token) Then
                Application.Exit()
                Environment.Exit(0)
            End If
            If _downloadCts IsNot Nothing Then
                _downloadCts.Dispose()
                _downloadCts = Nothing
            End If
            If Not Me.IsDisposed Then
                If prgDownload IsNot Nothing Then prgDownload.Visible = False
                If btnCancelDownload IsNot Nothing Then btnCancelDownload.Visible = False
                btnUpdateNow.Enabled = True
                btnUpdateNow.Text = "&Update && Restart Now"
            End If
        End Sub

        ' Offline fallback changelog and pre-2.1 history.
        Private Shared Function GetLocalHistory() As List(Of LocalHistoryEntry)
            Dim history As New List(Of LocalHistoryEntry)()

            history.Add(New LocalHistoryEntry With {
                .VersionText = "2.5",
                .ReleaseDate = New DateTime(2026, 10, 3),
                .Bullets = New List(Of String) From {
                    "Details Window Stability: Fixed flickering, ghosting, and black leftovers when resizing or maximizing, especially on the Document Overview tab.",
                    "Truthful Roadmap: The step-by-step custody roadmap now shows where a document really is; freshly registered documents start at their assigned desk.",
                    "Real Timestamps Only: The roadmap and routing slip no longer invent completion dates; dates appear as actions truly happen.",
                    "Offline Routing Integrity: Routing a document offline now updates its status together with the custody log, matching the connected flow."
                }
            })

            history.Add(New LocalHistoryEntry With {
                .VersionText = "2.4",
                .ReleaseDate = New DateTime(2026, 10, 3),
                .Bullets = New List(Of String) From {
                    "Sidebar Navigation: Relocated What's New button to the bottom of the left navigation pane for easy access.",
                    "Top Header: Cleaned up the top banner to keep focus on badge scanning and user logout.",
                    "Clean Dialog Header: Removed circular badge graphic for a clean, distraction-free update window."
                }
            })

            history.Add(New LocalHistoryEntry With {
                .VersionText = "2.3",
                .ReleaseDate = New DateTime(2026, 10, 3),
                .Bullets = New List(Of String) From {
                    "Sharp Desktop Icons: Integrated 23 clean vector icons that stay sharp on all monitor display scalings.",
                    "Theme Contrast Tinting: Icons adapt automatically to match screen theme colors for high readability.",
                    "Scanned Attachments Folder: Station option to centralize PDF document scans into a shared drive folder.",
                    "Performance Tuning: Reduced workstation memory usage and improved application startup speed."
                }
            })

            history.Add(New LocalHistoryEntry With {
                .VersionText = "2.1",
                .ReleaseDate = New DateTime(2026, 10, 3),
                .Bullets = New List(Of String) From {
                    "What's New & Release Changelogs: Integrated release history and update notification window.",
                    "1-Click System Updates: Self-service update checker querying GitHub Releases with automated in-place restart.",
                    "Embedded Assembly Metadata: Official Bangsamoro Transition Authority metadata headers to prevent heuristic antivirus flagging.",
                    "Automated Stream Unblocking: Updated binaries automatically strip Windows Mark-of-the-Web zone identifiers.",
                    "Audit Trail Hash Chain: Cryptographic SHA-256 seal chains for tamper-evident document action logging.",
                    "Multi-Workstation Sync: Background peer heartbeat discovery and non-blocking SQL Server mirror merges."
                }
            })

            history.Add(New LocalHistoryEntry With {
                .VersionText = "2.0.0",
                .HeaderText = "v2.0 - September 2026",
                .Bullets = New List(Of String) From {
                    "External Intake Web Portal: Citizen and ministry intake portal with OTP verification and tracking.",
                    "Routing Slip Print Engine: Standard Bangsamoro parliamentary document routing slip print service.",
                    "Executive Action Directives: Secretary-General revision loops and status routing.",
                    "Controlled Storage Tracking: Cabinet, shelf, and box physical archive movement tracking.",
                    "High-DPI Civic Calm Theme: Accessible, high-contrast desktop visual design."
                }
            })

            history.Add(New LocalHistoryEntry With {
                .VersionText = "1.5.0",
                .HeaderText = "v1.5 - August 2026",
                .Bullets = New List(Of String) From {
                    "RFID Dual-Factor Security: Contactless smart card authentication with terminal auto-lockout defense.",
                    "Offline Outbox: Background sync to SQL Server with automatic offline cache fallback.",
                    "Role-Based Access Control: Granular permissions for Records Officers, Legal Counsel, and Office of the SG."
                }
            })

            history.Add(New LocalHistoryEntry With {
                .VersionText = "1.0.0",
                .HeaderText = "v1.0 - July 2026",
                .Bullets = New List(Of String) From {
                    "Document Tracking Engine: Central database tracking for communications, bills, vouchers, and travel orders.",
                    "Sequential Tracking Numbers: Automated document tracking code generator (COMM, LEG, FIN, TO).",
                    "OSG Desk Routing: Digital document routing and status transitions between division desks."
                }
            })

            Return history
        End Function

        ' Pad missing version parts with zero so "2.1" and "2.1.0.0" produce the same key.
        Friend Shared Function NormalizeVersion(tagOrVersion As String) As String
            Dim normalized = If(tagOrVersion, "").Trim().TrimStart("v"c, "V"c)
            Dim parsed As Version = Nothing
            If Version.TryParse(normalized, parsed) Then
                Return parsed.Major & "." & parsed.Minor & "." & Math.Max(parsed.Build, 0) & "." & Math.Max(parsed.Revision, 0)
            End If
            Return normalized
        End Function

        Private Shared Function FormatReleaseHeading(name As String, releasedAt As DateTime) As String
            If releasedAt = DateTime.MinValue Then Return name
            Return name & " - " & releasedAt.ToString("MMMM d, yyyy", CultureInfo.InvariantCulture)
        End Function

        Private Shared Function CuratedHeading(entry As LocalHistoryEntry) As String
            If entry.ReleaseDate <> DateTime.MinValue Then
                Return FormatReleaseHeading("v" & entry.VersionText, entry.ReleaseDate)
            End If
            Return entry.HeaderText
        End Function

        Private Class ChangelogSection
            Public Property Heading As String = ""
            Public Property SectionDate As DateTime = DateTime.MinValue
            Public Property Bullets As List(Of String)
            Public Property MarkdownBody As String = ""
        End Class

        Private Sub PopulateChangelog(remoteNotes As String, pastReleases As List(Of ReleaseHistoryEntry))
            Dim scrollBefore As Point = If(pnlGroups IsNot Nothing AndAlso pnlGroups.IsHandleCreated,
                                           pnlGroups.AutoScrollPosition, New Point(0, 0))
            If pnlGroups Is Nothing Then Return
            SetRedraw(pnlGroups, False)
            For Each v In _groupViews
                v.Panel.Dispose()
            Next
            _groupViews.Clear()
            pnlGroups.Controls.Clear()

            Dim sections = BuildSections(remoteNotes, pastReleases)
            Dim groups = GroupSectionsByMonth(sections)

            Dim bodyWidth = GroupBodyWidth()
            For Each group In groups
                Dim view = BuildMonthGroupView(group, bodyWidth)
                _groupViews.Add(view)
                pnlGroups.Controls.Add(view.Panel)
            Next

            SetRedraw(pnlGroups, True)
            Try
                pnlGroups.AutoScrollPosition = New Point(-scrollBefore.X, -scrollBefore.Y)
            Catch
            End Try
        End Sub

        Private Function BuildSections(remoteNotes As String, pastReleases As List(Of ReleaseHistoryEntry)) As List(Of ChangelogSection)
            Dim sections As New List(Of ChangelogSection)()

            If Not String.IsNullOrWhiteSpace(remoteNotes) Then
                sections.Add(New ChangelogSection With {
                    .Heading = FormatReleaseHeading($"Latest Release Notes ({AppUpdateService.FormatTagLabel(_latestVersion)})", _latestPublishedAt),
                    .SectionDate = _latestPublishedAt,
                    .MarkdownBody = remoteNotes.Trim()
                })
            End If

            Dim shownVersions As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            If pastReleases IsNot Nothing Then
                For Each rel In pastReleases
                    Dim versionKey = NormalizeVersion(rel.TagName)
                    If versionKey = NormalizeVersion(_latestVersion) AndAlso Not String.IsNullOrWhiteSpace(remoteNotes) Then
                        shownVersions.Add(versionKey)
                        Continue For
                    End If

                    Dim localEntry = GetLocalHistory().FirstOrDefault(Function(h) NormalizeVersion(h.VersionText) = versionKey)
                    If localEntry IsNot Nothing Then
                        Dim headingDate = If(rel.PublishedAt <> DateTime.MinValue, rel.PublishedAt, localEntry.ReleaseDate)
                        sections.Add(New ChangelogSection With {
                            .Heading = FormatReleaseHeading(AppUpdateService.FormatTagLabel(rel.TagName), headingDate),
                            .SectionDate = headingDate,
                            .Bullets = localEntry.Bullets
                        })
                    Else
                        sections.Add(New ChangelogSection With {
                            .Heading = FormatReleaseHeading(AppUpdateService.FormatTagLabel(rel.TagName), rel.PublishedAt),
                            .SectionDate = rel.PublishedAt,
                            .MarkdownBody = rel.Body.Trim()
                        })
                    End If
                    shownVersions.Add(versionKey)
                Next
            End If

            For Each entry In GetLocalHistory()
                If shownVersions.Contains(NormalizeVersion(entry.VersionText)) Then Continue For
                sections.Add(New ChangelogSection With {
                    .Heading = CuratedHeading(entry),
                    .SectionDate = entry.ReleaseDate,
                    .Bullets = entry.Bullets
                })
            Next

            Return sections
        End Function

        Private Shared Function GroupSectionsByMonth(sections As List(Of ChangelogSection)) As List(Of MonthGroup)
            Dim groups As New List(Of MonthGroup)()
            Dim groupIndex As New Dictionary(Of String, MonthGroup)()
            Dim undated As New List(Of ChangelogSection)()
            For Each sec In sections
                Dim monthKey = SectionMonthKey(sec)
                If monthKey = "" Then
                    undated.Add(sec)
                    Continue For
                End If
                Dim group As MonthGroup = Nothing
                If Not groupIndex.TryGetValue(monthKey, group) Then
                    Dim year = CInt(monthKey.Substring(0, 4))
                    Dim month = CInt(monthKey.Substring(5, 2))
                    Dim firstDate = If(sec.SectionDate <> DateTime.MinValue, sec.SectionDate, New DateTime(year, month, 1))
                    group = New MonthGroup With {
                        .Title = firstDate.ToString("MMMM yyyy", CultureInfo.InvariantCulture),
                        .Expanded = IsCurrentMonth(firstDate)
                    }
                    groupIndex.Add(monthKey, group)
                    groups.Add(group)
                End If
                group.Sections.Add(sec)
            Next
            If undated.Count > 0 Then
                Dim bucket As New MonthGroup With {.Title = "Earlier releases", .Expanded = False}
                For Each sec In undated
                    bucket.Sections.Add(sec)
                Next
                groups.Add(bucket)
            End If
            Return groups
        End Function

        Private Shared Function SectionMonthKey(sec As ChangelogSection) As String
            Dim d As DateTime = sec.SectionDate
            If d = DateTime.MinValue Then
                Dim idx = sec.Heading.IndexOf(" - ", StringComparison.Ordinal)
                If idx < 0 Then Return ""
                If Not Date.TryParse(sec.Heading.Substring(idx + 3), CultureInfo.InvariantCulture, DateTimeStyles.None, d) Then Return ""
            End If
            Return d.Year.ToString("0000") & "-" & d.Month.ToString("00")
        End Function

        Private Shared Function IsCurrentMonth(d As DateTime) As Boolean
            Return d <> DateTime.MinValue AndAlso d.Year = Date.Now.Year AndAlso d.Month = Date.Now.Month
        End Function

        Private Shared Function GroupHeaderText(g As MonthGroup) As String
            Dim chevron = If(g.Expanded, ChrW(&H25BE), ChrW(&H25B8))
            Dim count = g.Sections.Count
            Return chevron & "  " & g.Title & "   ·   " & count & If(count = 1, " release", " releases")
        End Function

        Private Shared Function GroupHeight(view As MonthGroupView) As Integer
            If Not view.Group.Expanded Then Return GroupHeaderHeight
            Return GroupHeaderHeight + GroupBodyTopOffset + view.Body.Height + 4
        End Function

        Private Function BuildMonthGroupView(g As MonthGroup, bodyWidth As Integer) As MonthGroupView
            Dim view As New MonthGroupView With {.Group = g}
            Dim innerWidth = Math.Max(180, bodyWidth - (GroupBodyLeftOffset * 2))

            view.Body = New RichTextBox With {
                .Width = innerWidth,
                .ReadOnly = True,
                .BorderStyle = BorderStyle.None,
                .BackColor = CivicCalmTheme.ColorSurface,
                .ForeColor = CivicCalmTheme.ColorInk,
                .Font = FontChangelogBody,
                .ScrollBars = RichTextBoxScrollBars.None,
                .TabStop = False,
                .DetectUrls = False,
                .Visible = True,
                .Height = 10
            }
            AddHandler view.Body.ContentsResized, Sub(s, e)
                                                       view.Body.Height = CInt(e.NewRectangle.Height) + 4
                                                       If view.Panel IsNot Nothing Then view.Panel.Height = GroupHeight(view)
                                                   End Sub

            ' Create handle while visible so RichEdit formats and reports ContentsResized immediately.
            Dim forced = view.Body.Handle
            RenderSectionsInto(view.Body, g)
            view.Body.Visible = g.Expanded

            view.Header = New LinkLabel With {
                .Width = bodyWidth,
                .Text = GroupHeaderText(g),
                .Height = GroupHeaderHeight,
                .TextAlign = ContentAlignment.MiddleLeft,
                .LinkBehavior = LinkBehavior.NeverUnderline,
                .BackColor = CivicCalmTheme.ColorWell,
                .LinkColor = CivicCalmTheme.ColorInk,
                .ActiveLinkColor = CivicCalmTheme.ColorPrimary,
                .Cursor = Cursors.Hand,
                .Padding = New Padding(GroupBodyLeftOffset, 0, GroupBodyLeftOffset, 0),
                .Font = CivicCalmTheme.FontSectionHeader,
                .AutoSize = False
            }
            AddHandler view.Header.LinkClicked, Sub(s, e)
                                                     g.Expanded = Not g.Expanded
                                                     view.Body.Visible = g.Expanded
                                                     view.Header.Text = GroupHeaderText(g)
                                                     view.Panel.Height = GroupHeight(view)
                                                 End Sub

            view.Panel = New Panel With {
                .Width = bodyWidth,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Margin = New Padding(0, 0, 0, 10),
                .Height = GroupHeight(view)
            }
            view.Body.Location = New Point(GroupBodyLeftOffset, GroupHeaderHeight + GroupBodyTopOffset)
            view.Header.Location = New Point(0, 0)
            view.Panel.Controls.Add(view.Body)
            view.Panel.Controls.Add(view.Header)
            Return view
        End Function

        Private Sub RenderSectionsInto(rtb As RichTextBox, g As MonthGroup)
            Dim first As Boolean = True
            For Each sec In g.Sections
                If Not first Then AppendSpacerLine(rtb, FontChangelogSectionGap)
                AppendHeader(rtb, sec.Heading)
                If sec.Bullets IsNot Nothing Then
                    For Each b In sec.Bullets
                        AppendBullet(rtb, b)
                    Next
                Else
                    AppendMarkdownBody(rtb, sec.MarkdownBody)
                End If
                first = False
            Next
        End Sub

        Private Function GroupBodyWidth() As Integer
            If pnlGroups IsNot Nothing AndAlso pnlGroups.ClientSize.Width > 0 Then
                Dim w = pnlGroups.ClientSize.Width - pnlGroups.Padding.Horizontal - SystemInformation.VerticalScrollBarWidth
                If w >= 200 Then Return w
            End If
            Return 560
        End Function

        Private Const WmSetRedraw As Integer = &HB

        <System.Runtime.InteropServices.DllImport("user32.dll")>
        Private Shared Function SendMessage(hWnd As IntPtr, msg As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr
        End Function

        Private Sub SetRedraw(target As Control, enabled As Boolean)
            If target Is Nothing OrElse Not target.IsHandleCreated Then Return
            SendMessage(target.Handle, WmSetRedraw, If(enabled, New IntPtr(1), IntPtr.Zero), IntPtr.Zero)
            If enabled Then target.Invalidate()
        End Sub

        ' Formats the appended paragraph by exact character length so indents and fonts never bleed into subsequent paragraphs.
        Private Shared Sub AppendFormattedParagraph(rtb As RichTextBox, text As String, font As Font, color As Color, indent As Integer, hangingIndent As Integer, trailingSpacerFont As Font)
            rtb.Select(rtb.TextLength, 0)
            Dim startIdx = rtb.TextLength
            rtb.AppendText(text & vbCrLf)
            rtb.Select(startIdx, text.Length)
            rtb.SelectionFont = font
            rtb.SelectionColor = color
            rtb.SelectionIndent = indent
            rtb.SelectionHangingIndent = hangingIndent

            rtb.Select(rtb.TextLength, 0)
            rtb.SelectionIndent = 0
            rtb.SelectionHangingIndent = 0

            If trailingSpacerFont IsNot Nothing Then
                AppendSpacerLine(rtb, trailingSpacerFont)
            End If
        End Sub

        Private Shared Sub AppendSpacerLine(rtb As RichTextBox, spacerFont As Font)
            rtb.Select(rtb.TextLength, 0)
            Dim startIdx = rtb.TextLength
            rtb.AppendText(vbCrLf)
            rtb.Select(startIdx, rtb.TextLength - startIdx)
            rtb.SelectionFont = spacerFont
            rtb.SelectionIndent = 0
            rtb.SelectionHangingIndent = 0
            rtb.Select(rtb.TextLength, 0)
        End Sub

        Private Shared Sub AppendHeader(rtb As RichTextBox, text As String)
            AppendFormattedParagraph(rtb, text, FontChangelogHeader, CivicCalmTheme.ColorPrimary, 0, 0, FontChangelogSpacer)
        End Sub

        Private Shared Sub AppendSubheader(rtb As RichTextBox, text As String)
            AppendFormattedParagraph(rtb, text, FontChangelogSubheader, CivicCalmTheme.ColorInk, 0, 0, FontChangelogSpacer)
        End Sub

        Private Shared Sub AppendBullet(rtb As RichTextBox, text As String)
            AppendFormattedParagraph(rtb, "•  " & text, FontChangelogBody, CivicCalmTheme.ColorInk, 0, 16, FontChangelogSpacer)
        End Sub

        Private Shared Sub AppendBody(rtb As RichTextBox, text As String)
            AppendFormattedParagraph(rtb, text, FontChangelogBody, CivicCalmTheme.ColorInk, 0, 0, FontChangelogSpacer)
        End Sub

        ' Joins hard-wrapped Markdown lines into full paragraphs/bullets and renders #### category subheadings.
        Friend Shared Sub AppendMarkdownBody(rtb As RichTextBox, text As String)
            If String.IsNullOrWhiteSpace(text) Then Return

            Dim pendingBullet As String = Nothing
            Dim pendingParagraph As String = Nothing

            Dim flushPending As Action = Sub()
                                             If pendingBullet IsNot Nothing Then
                                                 AppendBullet(rtb, StripMarkdownEmphasis(pendingBullet))
                                                 pendingBullet = Nothing
                                             End If
                                             If pendingParagraph IsNot Nothing Then
                                                 AppendBody(rtb, StripMarkdownEmphasis(pendingParagraph))
                                                 pendingParagraph = Nothing
                                             End If
                                         End Sub

            For Each rawLine In text.Replace(vbCr, "").Split(New Char() {ControlChars.Lf})
                Dim line = rawLine.Trim()
                If line.Length = 0 Then
                    flushPending()
                    Continue For
                End If
                If line.StartsWith("<!--", StringComparison.Ordinal) Then
                    flushPending()
                    Continue For
                End If
                If line.StartsWith("#"c) Then
                    flushPending()
                    Dim heading = line.TrimStart("#"c).Trim()
                    ' Drop the top-level "What's New in vX.Y" title (already in the release header); keep section subheadings.
                    If heading.Length > 0 AndAlso Not heading.StartsWith("What's New", StringComparison.OrdinalIgnoreCase) Then
                        AppendSubheader(rtb, StripMarkdownEmphasis(heading))
                    End If
                    Continue For
                End If
                If line.StartsWith("- ", StringComparison.Ordinal) OrElse line.StartsWith("* ", StringComparison.Ordinal) Then
                    flushPending()
                    pendingBullet = line.Substring(2).Trim()
                ElseIf pendingBullet IsNot Nothing Then
                    pendingBullet &= " " & line
                ElseIf pendingParagraph IsNot Nothing Then
                    pendingParagraph &= " " & line
                Else
                    pendingParagraph = line
                End If
            Next

            flushPending()
        End Sub

        Private Shared Function StripMarkdownEmphasis(text As String) As String
            Return text.Replace("**", "").Replace("__", "")
        End Function
    End Class
End Namespace
