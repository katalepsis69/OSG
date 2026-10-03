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
            Public Property HeaderText As String = ""
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

        Private _hasUpdate As Boolean = False
        Private _latestVersion As String = ""
        Private _downloadUrl As String = ""
        Private _remoteNotes As String = ""
        Private _latestPublishedAt As DateTime = DateTime.MinValue
        Private _alreadyChecked As Boolean = False

        Public Sub New(Optional hasUpdate As Boolean = False, Optional latestVersion As String = "", Optional downloadUrl As String = "", Optional releaseNotes As String = "", Optional latestPublishedAt As DateTime = Nothing, Optional alreadyChecked As Boolean = False)
            _hasUpdate = hasUpdate
            _latestVersion = latestVersion
            _downloadUrl = downloadUrl
            _remoteNotes = releaseNotes
            _latestPublishedAt = latestPublishedAt
            _alreadyChecked = alreadyChecked

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

            ' Top Header Panel
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

            lblHeaderTitle = New Label With {
                .Text = "What's New in BTA OSG",
                .Font = CivicCalmTheme.FontFormTitle,
                .ForeColor = CivicCalmTheme.ColorInk,
                .AutoSize = True,
                .Location = New Point(20, 16)
            }
            pnlHeader.Controls.Add(lblHeaderTitle)

            Dim currentVer = Assembly.GetExecutingAssembly().GetName().Version
            Dim verStr = If(currentVer IsNot Nothing, $"v{AppUpdateService.FormatVersion(currentVer)}", "v?")

            lblHeaderSubtitle = New Label With {
                .Text = $"Current Version: {verStr}  -  Parliamentary Document Tracking System",
                .Font = CivicCalmTheme.FontMicrocopy,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .AutoSize = True,
                .Location = New Point(20, 42)
            }
            pnlHeader.Controls.Add(lblHeaderSubtitle)

            ' Status Badge
            lblStatusBadge = New Label With {
                .AutoSize = False,
                .Size = New Size(190, 28),
                .TextAlign = ContentAlignment.MiddleCenter,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Location = New Point(Math.Max(10, Me.ClientSize.Width - 210), 24)
            }

            ' Reposition badge relative to header client width to prevent docking offset bugs
            AddHandler pnlHeader.Resize, Sub(s, e)
                If lblStatusBadge IsNot Nothing Then
                    lblStatusBadge.Location = New Point(Math.Max(10, pnlHeader.ClientSize.Width - lblStatusBadge.Width - 20), 24)
                End If
            End Sub

            If _hasUpdate Then
                lblStatusBadge.Text = $"Update Available ({AppUpdateService.FormatTagLabel(_latestVersion)})"
                lblStatusBadge.BackColor = CivicCalmTheme.ColorStatusReceivedBg
                lblStatusBadge.ForeColor = CivicCalmTheme.ColorStatusReceivedFg
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
            pnlHeader.Controls.Add(lblStatusBadge)

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

            btnUpdateNow = New Button With {
                .Text = " &Update && Restart Now",
                .Size = New Size(205, 34),
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .BackColor = CivicCalmTheme.ColorPrimary,
                .ForeColor = Color.White,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(6, 0, 0, 0),
                .Image = AppAssets.GetIcon("arrow-clockwise", 16, Color.White),
                .ImageAlign = ContentAlignment.MiddleLeft,
                .TextAlign = ContentAlignment.MiddleCenter,
                .TextImageRelation = TextImageRelation.ImageBeforeText,
                .Padding = New Padding(8, 0, 8, 0)
            }
            btnUpdateNow.FlatAppearance.BorderSize = 0
            AddHandler btnUpdateNow.Click, Async Sub()
                btnUpdateNow.Enabled = False
                btnUpdateNow.Text = " Downloading..."
                Await AppUpdateService.DownloadAndApplyAsync(_downloadUrl, Me)
                If Not Me.IsDisposed Then
                    btnUpdateNow.Enabled = True
                    btnUpdateNow.Text = " &Update && Restart Now"
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

            history.Add(New LocalHistoryEntry With {
                .VersionText = "2.5",
                .HeaderText = "v2.5 - October 2026",
                .Bullets = New List(Of String) From {
                    "Details Window Stability: Fixed flickering, ghosting, and black leftovers when resizing or maximizing, especially on the Document Overview tab.",
                    "Truthful Roadmap: The step-by-step custody roadmap now shows where a document really is; freshly registered documents start at their assigned desk.",
                    "Real Timestamps Only: The roadmap and routing slip no longer invent completion dates; dates appear as actions truly happen.",
                    "Offline Routing Integrity: Routing a document offline now updates its status together with the custody log, matching the connected flow."
                }
            })

            history.Add(New LocalHistoryEntry With {
                .VersionText = "2.4",
                .HeaderText = "v2.4 - October 2026",
                .Bullets = New List(Of String) From {
                    "Sidebar Navigation: Relocated What's New button to the bottom of the left navigation pane for easy access.",
                    "Header Streamlining: Cleaned up the top banner to keep focus on badge scanning and user logout.",
                    "Clean Dialog Header: Removed circular badge graphic for a clean, distraction-free update window."
                }
            })

            history.Add(New LocalHistoryEntry With {
                .VersionText = "2.3",
                .HeaderText = "v2.3 - October 2026",
                .Bullets = New List(Of String) From {
                    "Sharp Desktop Icons: Integrated 23 clean vector icons that stay sharp on all monitor display scalings.",
                    "Theme Contrast Tinting: Icons adapt automatically to match screen theme colors for high readability.",
                    "Scanned Attachments Folder: Station option to centralize PDF document scans into a shared drive folder.",
                    "Performance Tuning: Reduced workstation memory usage and improved application startup speed."
                }
            })

            history.Add(New LocalHistoryEntry With {
                .VersionText = "2.1",
                .HeaderText = "v2.1 - October 2026",
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

        Private Shared Function NormalizeVersion(tagOrVersion As String) As String
            Dim normalized = If(tagOrVersion, "").Trim().TrimStart("v"c, "V"c)
            Dim parsed As Version = Nothing
            If Version.TryParse(normalized, parsed) Then
                Return parsed.ToString(3)
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

        Private Sub PopulateChangelog(remoteNotes As String, pastReleases As List(Of ReleaseHistoryEntry))
            rtbChangelog.Clear()

            ' If remote release notes are available from GitHub, show them at top
            If Not String.IsNullOrWhiteSpace(remoteNotes) Then
                AppendHeader(FormatReleaseHeading($"Latest Release Notes ({AppUpdateService.FormatTagLabel(_latestVersion)})", _latestPublishedAt))
                AppendMarkdownBody(remoteNotes.Trim())
                AppendSeparator()
            End If

            ' Past GitHub releases, newest first. The releases list is the single changelog
            ' source, so a release can never vanish between "latest" and the local list again
            ' (2.1.4 was skipped exactly that way). Where a curated local entry exists it is
            ' preferred; otherwise the release body renders as-is.
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
                        ' Curated entries carry their own dated two-part heading
                        ' ("v2.1 - October 2026"), matching the GitHub-fed entries.
                        AppendHeader(localEntry.HeaderText)
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
                AppendHeader(entry.HeaderText)
                For Each b In entry.Bullets
                    AppendBullet(b)
                Next
                AppendSeparator()
            Next

            rtbChangelog.SelectionStart = 0
            rtbChangelog.ScrollToCaret()
        End Sub

        Private Sub AppendHeader(text As String)
            rtbChangelog.SelectionFont = CivicCalmTheme.FontSectionHeader
            rtbChangelog.SelectionColor = CivicCalmTheme.ColorPrimary
            rtbChangelog.AppendText(text & vbCrLf)
        End Sub

        Private Sub AppendBullet(text As String)
            rtbChangelog.SelectionFont = CivicCalmTheme.FontBody
            rtbChangelog.SelectionColor = CivicCalmTheme.ColorInk
            rtbChangelog.AppendText("  •  " & text & vbCrLf)
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
