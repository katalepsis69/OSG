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
    ''' <summary>
    ''' Displays release history, recent changelogs, and 1-click update status.
    ''' Adheres to Civic Calm design system and WCAG AA contrast standards.
    ''' </summary>
    Public Class FormWhatsNew
        Inherits Form

        Private Class LocalHistoryEntry
            Public Property VersionText As String = ""
            ' Only for releases that predate tagging: a curated heading when no real date
            ' exists to format.
            Public Property HeaderText As String = ""
            Public Property ReleaseDate As DateTime = DateTime.MinValue
            Public Property Bullets As New List(Of String)()
        End Class

        Private lblHeaderTitle As Label
        Private lblHeaderSubtitle As Label
        Private lblStatusBadge As Label
        Private pnlGroups As FlowLayoutPanel
        Private pnlReading As Panel
        Private btnUpdateNow As Button
        Private btnClose As Button
        Private pnlHeader As Panel
        Private pnlFooter As Panel
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

        ' Month-group disclosure: every month is its own collapsible block with a live
        ' header row. The current month starts expanded, every earlier month starts
        ' collapsed (header row only), and each expands or collapses independently:
        ' collapsing one month never touches another.
        Private _pastReleases As List(Of ReleaseHistoryEntry)
        Private ReadOnly _groupViews As New List(Of MonthGroupView)()

        ' Reading surface: the changelog is long-form text, so it runs a step above the
        ' 9pt control body font with a matching header weight. Defined here once: this
        ' window's reading decision, deliberately not a new theme-wide token.
        Private Shared ReadOnly FontChangelogBody As New Font("Segoe UI", 9.75F, FontStyle.Regular)
        Private Shared ReadOnly FontChangelogHeader As New Font("Segoe UI", 10.5F, FontStyle.Bold)

        Private Const GroupHeaderHeight As Integer = 34

        ''' <summary>One month (or the undated "Earlier releases" bucket) and its sections.
        ''' Expanded defaults to the current calendar month; toggling is per group.</summary>
        Private Class MonthGroup
            Public Property Title As String = ""
            Public Property RepresentativeDate As DateTime = DateTime.MinValue
            Public Property Expanded As Boolean = False
            Public ReadOnly Sections As New List(Of ChangelogSection)()
        End Class

        ''' <summary>Live controls for one month group; kept so toggles and resizes can
        ''' retune heights without rebuilding the whole list.</summary>
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

            ' Top header: a two-column layout panel so the text rows grow with DPI instead
            ' of colliding at fixed y offsets, and the badge right-aligns by anchoring.
            pnlHeader = New Panel With {
                .Dock = DockStyle.Top,
                .Height = 84,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(20, 14, 20, 10)
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
                .BackColor = CivicCalmTheme.ColorSurface
            }
            headerLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            headerLayout.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            headerLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            headerLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))

            lblHeaderTitle = New Label With {
                .Text = "What's New in BTA OSG",
                .Font = CivicCalmTheme.FontFormTitle,
                .ForeColor = CivicCalmTheme.ColorInk,
                .AutoSize = True,
                .Margin = New Padding(0, 0, 0, 2)
            }
            headerLayout.Controls.Add(lblHeaderTitle, 0, 0)

            Dim currentVer = Assembly.GetExecutingAssembly().GetName().Version
            Dim verStr = If(currentVer IsNot Nothing, $"v{AppUpdateService.FormatVersion(currentVer)}", "v?")

            lblHeaderSubtitle = New Label With {
                .Text = $"Current Version: {verStr}  -  Parliamentary Document Tracking System",
                .Font = CivicCalmTheme.FontMicrocopy,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .AutoSize = True,
                .Margin = New Padding(0)
            }
            headerLayout.Controls.Add(lblHeaderSubtitle, 0, 1)

            ' Status badge
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

            If _hasUpdate Then
                lblStatusBadge.Text = $"Update Available ({AppUpdateService.FormatTagLabel(_latestVersion)})"
                lblStatusBadge.BackColor = CivicCalmTheme.ColorStatusReceivedBg
                lblStatusBadge.ForeColor = CivicCalmTheme.ColorStatusReceivedFg
            ElseIf _alreadyChecked AndAlso Not _checkSucceeded Then
                ' A failed check is not a verified fact: the badge must not claim "up to
                ' date" when nothing was compared.
                lblStatusBadge.Text = "Update Check Unavailable"
                lblStatusBadge.BackColor = CivicCalmTheme.ColorStatusReceivedBg
                lblStatusBadge.ForeColor = CivicCalmTheme.ColorWarning
            ElseIf Not String.IsNullOrEmpty(_latestVersion) Then
                lblStatusBadge.Text = $"System Up to Date ({verStr})"
                lblStatusBadge.BackColor = CivicCalmTheme.ColorPrimarySoft
                lblStatusBadge.ForeColor = CivicCalmTheme.ColorPrimary
            ElseIf _alreadyChecked Then
                lblStatusBadge.Text = $"System Up to Date ({verStr})"
                lblStatusBadge.BackColor = CivicCalmTheme.ColorPrimarySoft
                lblStatusBadge.ForeColor = CivicCalmTheme.ColorPrimary
            Else
                lblStatusBadge.Text = "Checking updates..."
                lblStatusBadge.BackColor = CivicCalmTheme.ColorWell
                lblStatusBadge.ForeColor = CivicCalmTheme.ColorInkMuted
            End If

            ' Bottom Footer Panel
            pnlFooter = New Panel With {
                .Dock = DockStyle.Bottom,
                .Height = 60,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(16, 12, 16, 12)
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
                .BackColor = CivicCalmTheme.ColorSurface
            }

            btnClose = New Button With {
                .Text = "&Close",
                .Size = New Size(95, 34),
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontBody,
                .BackColor = CivicCalmTheme.ColorWell,
                .ForeColor = CivicCalmTheme.ColorInk,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(6, 0, 0, 0)
            }
            btnClose.FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder
            AddHandler btnClose.Click, Sub() Me.Close()
            flpFooter.Controls.Add(btnClose)
            ' Enter and Escape both close through the standard modal mechanisms; closing
            ' while a download runs cancels it (the FormClosing handler below), it never
            ' strands one. No other default verb: a stray Enter must not start a download.
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

            ' Changelog reading plate in the middle: one collapsible month group per month
            ' flows top-down inside a padded white panel so long lines never run
            ' edge-to-edge (RichTextBox itself ignores Padding); the plate scrolls when
            ' the expanded backlog outgrows the window.
            Dim pnlContent As New Panel With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(20, 16, 20, 8),
                .BackColor = CivicCalmTheme.ColorCanvas
            }

            pnlReading = New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(18, 14, 18, 10)
            }

            pnlGroups = New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.TopDown,
                .WrapContents = False,
                .AutoScroll = True,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(2, 2, 10, 8)
            }

            pnlReading.Controls.Add(pnlGroups)
            pnlContent.Controls.Add(pnlReading)

            Me.Controls.Add(pnlContent)
            Me.Controls.Add(pnlFooter)
            Me.Controls.Add(pnlHeader)
            pnlContent.BringToFront()
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

            ' Past releases come from the GitHub list whenever the network allows; the local
            ' curated history is the offline fallback, so every version stays visible either
            ' way (2.1.4 once vanished exactly between those two sources).
            Dim pastReleases = Await AppUpdateService.GetRecentReleasesAsync()
            If Me.IsDisposed Then Return
            _pastReleases = pastReleases
            PopulateChangelog(_remoteNotes, pastReleases)
        End Sub

        Private Sub ShowUpdateNowButton()
            If btnUpdateNow IsNot Nothing OrElse String.IsNullOrEmpty(_downloadUrl) Then Return

            ' A thin progress strip under the changelog: without it the download is one
            ' invisible operation and the window reads as frozen until the app exits.
            If pnlReading IsNot Nothing Then
                prgDownload = New ProgressBar With {
                    .Dock = DockStyle.Bottom,
                    .Height = 6,
                    .Visible = False
                }
                pnlReading.Controls.Add(prgDownload)
                pnlGroups.BringToFront()
            End If

            ' AutoSize: the progress line ("Downloading 18.4 of 43.5 MB (42%)") clips a fixed
            ' 205px button mid-download.
            btnUpdateNow = CivicCalmTheme.CreateIconButton("&Update && Restart Now", "arrow-clockwise", backColor:=CivicCalmTheme.ColorPrimary, margin:=New Padding(6, 0, 0, 0))
            AddHandler btnUpdateNow.Click, AddressOf OnUpdateNowClicked
            flpFooter.Controls.Add(btnUpdateNow)

            ' The download client deliberately carries no overall timeout, so the only
            ' escape from a stalled stream is this explicit cancel affordance.
            btnCancelDownload = CivicCalmTheme.CreateIconButton("&Stop Download", "x", 14, minHeight:=30, margin:=New Padding(6, 0, 0, 0), visible:=False)
            AddHandler btnCancelDownload.Click, Sub()
                                                    If _downloadCts IsNot Nothing Then _downloadCts.Cancel()
                                                End Sub
            flpFooter.Controls.Add(btnCancelDownload)
        End Sub

        ''' <summary>The whole update lifecycle on one named handler: start the download,
        ''' exit on the UI thread once the apply script is armed (True), and on cancel or
        ''' failure dispose the token even when this dialog was closed mid-download (the
        ''' FormClosing cancel), then restore the affordances only while a living dialog
        ''' still shows them. A failure has reported "Update failed: ..." through the
        ''' progress line; a cancel reported "Download canceled." the same way.</summary>
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
                ' The apply script is running: it waits for this process to exit, swaps
                ' the exe, and restarts. Exiting here is on the UI thread (the Await
                ' resumed inside this handler), which the old service-side Exit was not.
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

        ''' <summary>
        ''' The offline fallback changelog: every version that ever shipped, so the dialog is
        ''' never empty without network. Online, the GitHub releases list is the source of
        ''' truth and these entries are used where they exist (they are more detailed) and
        ''' for versions that predate publishing releases to GitHub.
        ''' </summary>
        Private Shared Function GetLocalHistory() As List(Of LocalHistoryEntry)
            Dim history As New List(Of LocalHistoryEntry)()

            ' Releases from v2.1 on carry their real tag date (git tags v2.1 through v2.7
            ' were all cut 2026-10-03), so every heading shows a dated day like the
            ' GitHub-fed entries do. The three entries below without a date predate
            ' tagging and keep their month heading; no day was ever recorded for them.
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
                    "Header Streamlining: Cleaned up the top banner to keep focus on badge scanning and user logout.",
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
                    "High-DPI Civic Calm Theme: Modern, accessible, high-contrast WinForms visual design."
                }
            })

            history.Add(New LocalHistoryEntry With {
                .VersionText = "1.5.0",
                .HeaderText = "v1.5 - August 2026",
                .Bullets = New List(Of String) From {
                    "RFID Dual-Factor Security: Contactless smart card authentication with terminal auto-lockout defense.",
                    "Offline Resilient Outbox: Continuous background sync to SQL Server with seamless offline cache fallback.",
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

        ''' <summary>
        ''' Canonical key for matching release tags to curated entries: missing version parts
        ''' read as zero, so "2.1" and "2.1.0.0" share a key. Never Version.ToString(n): it
        ''' throws for a version shorter than n parts ("2.1" has no third field to print).
        ''' </summary>
        Friend Shared Function NormalizeVersion(tagOrVersion As String) As String
            Dim normalized = If(tagOrVersion, "").Trim().TrimStart("v"c, "V"c)
            Dim parsed As Version = Nothing
            If Version.TryParse(normalized, parsed) Then
                Return parsed.Major & "." & parsed.Minor & "." & Math.Max(parsed.Build, 0) & "." & Math.Max(parsed.Revision, 0)
            End If
            Return normalized
        End Function

        ''' <summary>
        ''' A release heading with the day it shipped: "v2.1.5 - October 3, 2026".
        ''' Releases whose date is unknown keep the bare name.
        ''' </summary>
        Private Shared Function FormatReleaseHeading(name As String, releasedAt As DateTime) As String
            If releasedAt = DateTime.MinValue Then Return name
            Return name & " - " & releasedAt.ToString("MMMM d, yyyy", CultureInfo.InvariantCulture)
        End Function

        ''' <summary>Curated entry heading: the real date when one is recorded, else the
        ''' entry's own text (pre-tagging releases keep their month heading).</summary>
        Private Shared Function CuratedHeading(entry As LocalHistoryEntry) As String
            If entry.ReleaseDate <> DateTime.MinValue Then
                Return FormatReleaseHeading("v" & entry.VersionText, entry.ReleaseDate)
            End If
            Return entry.HeaderText
        End Function

        ''' <summary>One changelog block: heading, its real date, and the body source
        ''' (curated bullets win over a GitHub body, matching the historical flow).</summary>
        Private Class ChangelogSection
            Public Property Heading As String = ""
            Public Property SectionDate As DateTime = DateTime.MinValue
            Public Property Bullets As List(Of String)
            Public Property MarkdownBody As String = ""
        End Class

        Private Sub PopulateChangelog(remoteNotes As String, pastReleases As List(Of ReleaseHistoryEntry))
            ' Rebuilt every time the network answers, and the two passes (constructor,
            ' OnShown) land about a second apart. Freeze the redraw across the rebuild so
            ' the swap is one frame instead of a growing stack of partial controls.
            Dim scrollBefore As Point = If(pnlGroups IsNot Nothing AndAlso pnlGroups.IsHandleCreated,
                                           pnlGroups.AutoScrollPosition, New Point(0, 0))
            If pnlGroups Is Nothing Then Return
            SetRedraw(pnlGroups, False)
            For Each v In _groupViews
                v.Panel.Dispose()
            Next
            _groupViews.Clear()
            pnlGroups.Controls.Clear()

            ' Rebuild pipeline: collect the sections (remote latest, GitHub releases,
            ' curated fallback), fold them into month groups, then build one view per group.
            Dim sections = BuildSections(remoteNotes, pastReleases)
            Dim groups = GroupSectionsByMonth(sections)

            ' One width for every group: sampled before any tall panel arrives, because
            ' the scrollbar's own appearance shrinks the client area mid-build and would
            ' otherwise give later months a narrower wrap than earlier ones.
            Dim bodyWidth = GroupBodyWidth()
            For Each group In groups
                Dim view = BuildMonthGroupView(group, bodyWidth)
                _groupViews.Add(view)
                pnlGroups.Controls.Add(view.Panel)
            Next

            SetRedraw(pnlGroups, True)
            ' Hold the reader's place when the network answer rebuilds the list underneath.
            Try
                pnlGroups.AutoScrollPosition = New Point(-scrollBefore.X, -scrollBefore.Y)
            Catch
            End Try
        End Sub

        ''' <summary>Collects the changelog sections in display order: the remote latest
        ''' release body, then the GitHub releases list (curated bullets preferred, one
        ''' dated scheme for every entry), then the curated local history as the offline
        ''' fallback for versions the API does not cover.</summary>
        Private Function BuildSections(remoteNotes As String, pastReleases As List(Of ReleaseHistoryEntry)) As List(Of ChangelogSection)
            Dim sections As New List(Of ChangelogSection)()

            ' If remote release notes are available from GitHub, show them at top
            If Not String.IsNullOrWhiteSpace(remoteNotes) Then
                sections.Add(New ChangelogSection With {
                    .Heading = FormatReleaseHeading($"Latest Release Notes ({AppUpdateService.FormatTagLabel(_latestVersion)})", _latestPublishedAt),
                    .SectionDate = _latestPublishedAt,
                    .MarkdownBody = remoteNotes.Trim()
                })
            End If

            ' Past GitHub releases, newest first. The releases list is the single changelog
            ' source, so a release can never vanish between "latest" and the local list again
            ' (2.1.4 was skipped exactly that way). Where a curated local entry exists its
            ' bullets are preferred, but the heading still takes the GitHub published date:
            ' one dated scheme ("October 3, 2026") for every entry on screen.
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

            ' Local curated history: fills versions the API does not cover, and is the whole
            ' list when offline.
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

        ''' <summary>Folds flat sections into month groups, newest first. The current
        ''' month expands; every other month becomes its own collapsed header row that
        ''' opens and closes on click: never all months at once. Sections whose month
        ''' cannot be established fall into one trailing "Earlier releases" bucket
        ''' rather than a wrong month.</summary>
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
                        .RepresentativeDate = firstDate,
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

        ''' <summary>The grouping key for one section: its real date's month, else the
        ''' month parsed from curated month headings ("v1.0 - July 2026"; the day was
        ''' never recorded, but the month is real). No month at all returns "" and the
        ''' section lands in the trailing "Earlier releases" bucket.</summary>
        Private Shared Function SectionMonthKey(sec As ChangelogSection) As String
            Dim d As DateTime = sec.SectionDate
            If d = DateTime.MinValue Then
                Dim idx = sec.Heading.IndexOf(" - ", StringComparison.Ordinal)
                If idx < 0 Then Return ""
                If Not Date.TryParse(sec.Heading.Substring(idx + 3), CultureInfo.InvariantCulture, DateTimeStyles.None, d) Then Return ""
            End If
            Return d.Year.ToString("0000") & "-" & d.Month.ToString("00")
        End Function

        ''' <summary>A section dated in the current calendar month renders expanded; unknown
        ''' dates cannot be proven current, so they read as older (collapsed).</summary>
        Private Shared Function IsCurrentMonth(d As DateTime) As Boolean
            Return d <> DateTime.MinValue AndAlso d.Year = Date.Now.Year AndAlso d.Month = Date.Now.Month
        End Function

        Private Shared Function GroupHeaderText(g As MonthGroup) As String
            ' Geometric chevrons (not emoji): they carry the real open/closed state.
            Dim chevron = If(g.Expanded, ChrW(&H25BE), ChrW(&H25B8))
            Dim count = g.Sections.Count
            Return chevron & "  " & g.Title & "   ·   " & count & If(count = 1, " release", " releases")
        End Function

        Private Shared Function GroupHeight(view As MonthGroupView) As Integer
            Dim bodyHeight = If(view.Group.Expanded, view.Body.Height, 0)
            Return GroupHeaderHeight + bodyHeight + 4
        End Function

        ''' <summary>One month block: a clickable header row and its own read-only body.
        ''' The body is built at a fixed width so ContentsResized (which fires on text
        ''' change) sizes it during the build; forcing the handle makes that happen now,
        ''' not at first paint.</summary>
        Private Function BuildMonthGroupView(g As MonthGroup, bodyWidth As Integer) As MonthGroupView
            Dim view As New MonthGroupView With {.Group = g}

            view.Body = New RichTextBox With {
                .Width = bodyWidth,
                .ReadOnly = True,
                .BorderStyle = BorderStyle.None,
                .BackColor = CivicCalmTheme.ColorSurface,
                .ForeColor = CivicCalmTheme.ColorInk,
                .Font = FontChangelogBody,
                .ScrollBars = RichTextBoxScrollBars.None,
                .TabStop = False,
                .Visible = g.Expanded,
                .Height = 10
            }
            AddHandler view.Body.ContentsResized, Sub(s, e)
                                                       view.Body.Height = CInt(e.NewRectangle.Height) + 6
                                                       ' The panel does not exist yet when handle
                                                       ' creation fires this during the build.
                                                       If view.Panel IsNot Nothing Then view.Panel.Height = GroupHeight(view)
                                                   End Sub
            RenderSectionsInto(view.Body, g)
            Dim forced = view.Body.Handle ' force creation so ContentsResized runs now

            view.Header = New LinkLabel With {
                .Width = bodyWidth,
                .Text = GroupHeaderText(g),
                .Height = GroupHeaderHeight,
                .TextAlign = ContentAlignment.MiddleLeft,
                .LinkBehavior = LinkBehavior.HoverUnderline,
                .LinkColor = CivicCalmTheme.ColorPrimary,
                .ActiveLinkColor = CivicCalmTheme.ColorPrimary,
                .Padding = New Padding(0, 6, 0, 4),
                .Font = FontChangelogHeader,
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
                .Margin = New Padding(0, 0, 0, 14),
                .Height = GroupHeight(view)
            }
            ' Absolute layout inside the block: header row on top, body beneath it, both
            ' at the fixed width chosen above.
            view.Body.Location = New Point(0, GroupHeaderHeight)
            view.Header.Location = New Point(0, 0)
            view.Panel.Controls.Add(view.Body)
            view.Panel.Controls.Add(view.Header)
            Return view
        End Function

        Private Sub RenderSectionsInto(rtb As RichTextBox, g As MonthGroup)
            Dim first As Boolean = True
            For Each sec In g.Sections
                If Not first Then AppendSectionGap(rtb)
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

        ''' <summary>The fixed wrap width for month bodies. The window is a FixedDialog,
        ''' so widths never change after layout: bodies are built at this width once and
        ''' ContentsResized (reliable on text change) settles their heights. The scrollbar
        ''' width is reserved up front so the scrollbar's arrival can never reflow.</summary>
        Private Function GroupBodyWidth() As Integer
            If pnlGroups IsNot Nothing AndAlso pnlGroups.IsHandleCreated Then
                Dim w = pnlGroups.ClientSize.Width - pnlGroups.Padding.Horizontal - SystemInformation.VerticalScrollBarWidth
                If w >= 200 Then Return w
            End If
            Return 540
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

        ''' <summary>Two blank lines between rendered sections: the release blocks must
        ''' read as separate documents, not one continuous wall of bullets.</summary>
        Private Sub AppendSectionGap(rtb As RichTextBox)
            rtb.AppendText(vbCrLf & vbCrLf)
        End Sub

        Private Sub AppendHeader(rtb As RichTextBox, text As String)
            rtb.SelectionFont = FontChangelogHeader
            rtb.SelectionColor = CivicCalmTheme.ColorPrimary
            rtb.AppendText(text & vbCrLf)
        End Sub

        Private Sub AppendBullet(rtb As RichTextBox, text As String)
            Dim startIdx = rtb.TextLength
            rtb.SelectionFont = FontChangelogBody
            rtb.SelectionColor = CivicCalmTheme.ColorInk
            rtb.AppendText("•  " & text & vbCrLf)
            ' Hanging indent: wrapped lines align under the text, not under the bullet, so
            ' long changelog lines read as one item instead of a ragged block.
            rtb.Select(startIdx, rtb.TextLength - startIdx)
            rtb.SelectionIndent = 0
            rtb.SelectionHangingIndent = 18
            rtb.SelectionLength = 0
        End Sub

        ''' <summary>
        ''' GitHub release bodies arrive as markdown. Render the shape this project emits
        ''' (a heading, blank lines, dash bullets, bold prefixes) as styled text and drop
        ''' the markup, so "###" never reaches the eye. The body's own heading is dropped
        ''' because the section header above it already names the version.
        ''' </summary>
        Private Sub AppendMarkdownBody(rtb As RichTextBox, text As String)
            If String.IsNullOrWhiteSpace(text) Then Return

            For Each rawLine In text.Replace(vbCr, "").Split(New Char() {ControlChars.Lf})
                Dim line = rawLine.Trim()
                If line.Length = 0 Then Continue For
                If line.StartsWith("#"c) Then Continue For
                ' The release body carries an invisible build stamp for the updater; it is
                ' machinery, not changelog text.
                If line.StartsWith("<!--", StringComparison.Ordinal) Then Continue For

                If line.StartsWith("- ") OrElse line.StartsWith("* ") Then
                    AppendBullet(rtb, StripMarkdownEmphasis(line.Substring(2).Trim()))
                Else
                    AppendBody(rtb, StripMarkdownEmphasis(line))
                End If
            Next
        End Sub

        Private Shared Function StripMarkdownEmphasis(text As String) As String
            Return text.Replace("**", "").Replace("__", "")
        End Function

        Private Sub AppendBody(rtb As RichTextBox, text As String)
            rtb.SelectionFont = FontChangelogBody
            ' Body prose is primary content, so it reads in full ink; the muted tone is
            ' reserved for the header microcopy (contrast verified: both pass WCAG AA).
            rtb.SelectionColor = CivicCalmTheme.ColorInk
            rtb.AppendText(text & vbCrLf)
        End Sub
    End Class
End Namespace
