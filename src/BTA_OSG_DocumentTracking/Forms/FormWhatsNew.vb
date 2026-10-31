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
        Private rtbChangelog As RichTextBox
        Private btnUpdateNow As Button
        Private btnClose As Button
        Private pnlHeader As Panel
        Private pnlFooter As Panel
        Private flpFooter As FlowLayoutPanel
        Private prgDownload As ProgressBar

        Private _hasUpdate As Boolean = False
        Private _latestVersion As String = ""
        Private _downloadUrl As String = ""
        Private _remoteNotes As String = ""
        Private _latestPublishedAt As DateTime = DateTime.MinValue
        Private _alreadyChecked As Boolean = False
        Private _checkSucceeded As Boolean = True

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
            Me.Size = New Size(640, 540)
            Me.MinimumSize = New Size(520, 420)
            Me.StartPosition = FormStartPosition.CenterParent
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.ShowInTaskbar = False
            Me.BackColor = CivicCalmTheme.ColorCanvas
            Me.Font = CivicCalmTheme.FontBody
            Me.KeyPreview = True

            ' Escape key closes dialog
            AddHandler Me.KeyDown, Sub(s, e)
                If e.KeyCode = Keys.Escape Then Me.Close()
            End Sub

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

            If _hasUpdate AndAlso Not String.IsNullOrEmpty(_downloadUrl) Then
                ShowUpdateNowButton()
            End If

            pnlFooter.Controls.Add(flpFooter)

            ' Changelog RichTextBox in Middle
            Dim pnlContent As New Panel With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(20, 16, 20, 16),
                .BackColor = CivicCalmTheme.ColorCanvas
            }

            rtbChangelog = New RichTextBox With {
                .Dock = DockStyle.Fill,
                .ReadOnly = True,
                .BorderStyle = BorderStyle.None,
                .BackColor = CivicCalmTheme.ColorSurface,
                .ForeColor = CivicCalmTheme.ColorInk,
                .Font = CivicCalmTheme.FontBody,
                .ScrollBars = RichTextBoxScrollBars.Vertical,
                .Margin = New Padding(0)
            }
            pnlContent.Controls.Add(rtbChangelog)

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
            PopulateChangelog(_remoteNotes, pastReleases)
        End Sub

        Private Sub ShowUpdateNowButton()
            If btnUpdateNow IsNot Nothing OrElse String.IsNullOrEmpty(_downloadUrl) Then Return

            ' A thin progress strip under the changelog: without it the download is one
            ' invisible operation and the window reads as frozen until the app exits.
            Dim hostPanel = TryCast(rtbChangelog.Parent, Panel)
            If hostPanel IsNot Nothing Then
                prgDownload = New ProgressBar With {
                    .Dock = DockStyle.Bottom,
                    .Height = 6,
                    .Visible = False
                }
                hostPanel.Controls.Add(prgDownload)
                rtbChangelog.BringToFront()
            End If

            ' AutoSize: the progress line ("Downloading 18.4 of 43.5 MB (42%)") clips a fixed
            ' 205px button mid-download.
            btnUpdateNow = CivicCalmTheme.CreateIconButton("&Update && Restart Now", "arrow-clockwise", backColor:=CivicCalmTheme.ColorPrimary, margin:=New Padding(6, 0, 0, 0))
            AddHandler btnUpdateNow.Click, Async Sub()
                btnUpdateNow.Enabled = False
                If prgDownload IsNot Nothing Then prgDownload.Visible = True
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
                Await AppUpdateService.DownloadAndApplyAsync(_downloadUrl, Me, reporter)
                ' Success exits the process from the service; this restore is the failure path
                ' (the service has already shown the error dialog by here).
                If Not Me.IsDisposed Then
                    If prgDownload IsNot Nothing Then prgDownload.Visible = False
                    btnUpdateNow.Enabled = True
                    btnUpdateNow.Text = "&Update && Restart Now"
                End If
            End Sub
            flpFooter.Controls.Add(btnUpdateNow)
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

        Private Sub PopulateChangelog(remoteNotes As String, pastReleases As List(Of ReleaseHistoryEntry))
            ' Rebuilt every time the network answers, and the two passes (constructor,
            ' OnShown) land about a second apart. Freeze the redraw across the rebuild:
            ' an RTB repaints on every append and jumps its view to the insertion point,
            ' so without the freeze the open sequence is: stale list, auto-scroll to the
            ' bottom, snap back to the top, list again. One frame instead of four.
            Dim keepTop As Integer = 0
            If rtbChangelog.TextLength > 0 Then keepTop = Math.Max(0, rtbChangelog.GetCharIndexFromPosition(New Point(4, 4)))
            SetRedraw(False)
            rtbChangelog.Clear()

            ' If remote release notes are available from GitHub, show them at top
            If Not String.IsNullOrWhiteSpace(remoteNotes) Then
                AppendHeader(FormatReleaseHeading($"Latest Release Notes ({AppUpdateService.FormatTagLabel(_latestVersion)})", _latestPublishedAt))
                AppendMarkdownBody(remoteNotes.Trim())
                AppendSeparator()
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
                        AppendHeader(FormatReleaseHeading(AppUpdateService.FormatTagLabel(rel.TagName), headingDate))
                        For Each b In localEntry.Bullets
                            AppendBullet(b)
                        Next
                    Else
                        AppendHeader(FormatReleaseHeading(AppUpdateService.FormatTagLabel(rel.TagName), rel.PublishedAt))
                        AppendMarkdownBody(rel.Body.Trim())
                    End If
                    AppendSeparator()
                    shownVersions.Add(versionKey)
                Next
            End If

            ' Local curated history: fills versions the API does not cover, and is the whole
            ' list when offline.
            For Each entry In GetLocalHistory()
                If shownVersions.Contains(NormalizeVersion(entry.VersionText)) Then Continue For
                AppendHeader(CuratedHeading(entry))
                For Each b In entry.Bullets
                    AppendBullet(b)
                Next
                AppendSeparator()
            Next

            ' Land at the top on open, and hold the reader's place when the network answer
            ' rebuilds the list underneath them.
            rtbChangelog.SelectionStart = Math.Min(keepTop, rtbChangelog.TextLength)
            rtbChangelog.SelectionLength = 0
            rtbChangelog.ScrollToCaret()
            SetRedraw(True)
        End Sub

        Private Const WmSetRedraw As Integer = &HB

        <System.Runtime.InteropServices.DllImport("user32.dll")>
        Private Shared Function SendMessage(hWnd As IntPtr, msg As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr
        End Function

        Private Sub SetRedraw(enabled As Boolean)
            If rtbChangelog Is Nothing OrElse Not rtbChangelog.IsHandleCreated Then Return
            SendMessage(rtbChangelog.Handle, WmSetRedraw, If(enabled, New IntPtr(1), IntPtr.Zero), IntPtr.Zero)
            If enabled Then rtbChangelog.Invalidate()
        End Sub

        Private Sub AppendHeader(text As String)
            rtbChangelog.SelectionFont = CivicCalmTheme.FontSectionHeader
            rtbChangelog.SelectionColor = CivicCalmTheme.ColorPrimary
            rtbChangelog.AppendText(text & vbCrLf)
        End Sub

        Private Sub AppendBullet(text As String)
            Dim startIdx = rtbChangelog.TextLength
            rtbChangelog.SelectionFont = CivicCalmTheme.FontBody
            rtbChangelog.SelectionColor = CivicCalmTheme.ColorInk
            rtbChangelog.AppendText("•  " & text & vbCrLf)
            ' Hanging indent: wrapped lines align under the text, not under the bullet, so
            ' long changelog lines read as one item instead of a ragged block.
            rtbChangelog.Select(startIdx, rtbChangelog.TextLength - startIdx)
            rtbChangelog.SelectionIndent = 0
            rtbChangelog.SelectionHangingIndent = 16
            rtbChangelog.SelectionLength = 0
        End Sub

        ''' <summary>
        ''' GitHub release bodies arrive as markdown. Render the shape this project emits
        ''' (a heading, blank lines, dash bullets, bold prefixes) as styled text and drop
        ''' the markup, so "###" never reaches the eye. The body's own heading is dropped
        ''' because the section header above it already names the version.
        ''' </summary>
        Private Sub AppendMarkdownBody(text As String)
            If String.IsNullOrWhiteSpace(text) Then Return

            For Each rawLine In text.Replace(vbCr, "").Split(New Char() {ControlChars.Lf})
                Dim line = rawLine.Trim()
                If line.Length = 0 Then Continue For
                If line.StartsWith("#"c) Then Continue For
                ' The release body carries an invisible build stamp for the updater; it is
                ' machinery, not changelog text.
                If line.StartsWith("<!--", StringComparison.Ordinal) Then Continue For

                If line.StartsWith("- ") OrElse line.StartsWith("* ") Then
                    AppendBullet(StripMarkdownEmphasis(line.Substring(2).Trim()))
                Else
                    AppendBody(StripMarkdownEmphasis(line))
                End If
            Next
        End Sub

        Private Shared Function StripMarkdownEmphasis(text As String) As String
            Return text.Replace("**", "").Replace("__", "")
        End Function

        Private Sub AppendBody(text As String)
            rtbChangelog.SelectionFont = CivicCalmTheme.FontBody
            ' Body prose is primary content, so it reads in full ink; the muted tone is
            ' reserved for the header microcopy (contrast verified: both pass WCAG AA).
            rtbChangelog.SelectionColor = CivicCalmTheme.ColorInk
            rtbChangelog.AppendText(text & vbCrLf)
        End Sub

        Private Sub AppendSeparator()
            rtbChangelog.AppendText(vbCrLf)
        End Sub
    End Class
End Namespace
