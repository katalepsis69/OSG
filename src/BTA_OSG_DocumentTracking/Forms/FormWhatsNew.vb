Option Explicit On
Option Strict On

Imports System
Imports System.Drawing
Imports System.Reflection
Imports System.Windows.Forms

Namespace BTA_OSG
    ''' <summary>
    ''' Displays release history, recent changelogs, and 1-click update status.
    ''' Adheres to Civic Calm design system and WCAG AA contrast standards.
    ''' </summary>
    Public Class FormWhatsNew
        Inherits Form

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
        Private _alreadyChecked As Boolean = False

        Public Sub New(Optional hasUpdate As Boolean = False, Optional latestVersion As String = "", Optional downloadUrl As String = "", Optional releaseNotes As String = "", Optional alreadyChecked As Boolean = False)
            _hasUpdate = hasUpdate
            _latestVersion = latestVersion
            _downloadUrl = downloadUrl
            _remoteNotes = releaseNotes
            _alreadyChecked = alreadyChecked

            InitializeComponent()
            PopulateChangelog(_remoteNotes)
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
            Dim verStr = If(currentVer IsNot Nothing, $"v{currentVer.ToString(3)}", "v2.1.0")

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
                lblStatusBadge.Text = $"Update Available ({_latestVersion})"
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
                Dim verStr = If(currentVer IsNot Nothing, $"v{currentVer.ToString(3)}", "v2.1.0")

                If info.Success Then
                    _latestVersion = info.LatestVersion
                    _downloadUrl = info.DownloadUrl
                    _remoteNotes = info.ReleaseNotes
                    _hasUpdate = info.HasUpdate

                    If _hasUpdate Then
                        lblStatusBadge.Text = $"Update Available ({_latestVersion})"
                        lblStatusBadge.BackColor = CivicCalmTheme.ColorStatusReceivedBg
                        lblStatusBadge.ForeColor = CivicCalmTheme.ColorStatusReceivedFg
                        ShowUpdateNowButton()
                    Else
                        lblStatusBadge.Text = $"System Up to Date ({verStr})"
                        lblStatusBadge.BackColor = CivicCalmTheme.ColorPrimarySoft
                        lblStatusBadge.ForeColor = CivicCalmTheme.ColorPrimary
                    End If

                    PopulateChangelog(_remoteNotes)
                Else
                    lblStatusBadge.Text = $"System Ready ({verStr})"
                    lblStatusBadge.BackColor = CivicCalmTheme.ColorPrimarySoft
                    lblStatusBadge.ForeColor = CivicCalmTheme.ColorPrimary
                End If
            End If
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

        Private Sub PopulateChangelog(remoteNotes As String)
            rtbChangelog.Clear()

            ' If remote release notes are available from GitHub, show them at top
            If Not String.IsNullOrWhiteSpace(remoteNotes) Then
                AppendHeader($"Latest Release Notes ({_latestVersion})")
                AppendBody(remoteNotes.Trim())
                AppendSeparator()
            End If

            ' Built-in Version History Log
            AppendHeader("Version 2.1.3 - October 2026")
            AppendBullet("Sidebar Navigation: Relocated What's New button to the bottom of the left navigation pane for easy access.")
            AppendBullet("Header Streamlining: Cleaned up the top banner to keep focus on badge scanning and user logout.")
            AppendBullet("Clean Dialog Header: Removed circular badge graphic for a clean, distraction-free update window.")

            AppendSeparator()

            AppendHeader("Version 2.1.2 - October 2026")
            AppendBullet("Sharp Desktop Icons: Integrated 23 clean vector icons that stay sharp on all monitor display scalings.")
            AppendBullet("Theme Contrast Tinting: Icons adapt automatically to match screen theme colors for high readability.")
            AppendBullet("Scanned Attachments Folder: Station option to centralize PDF document scans into a shared drive folder.")
            AppendBullet("Performance Tuning: Reduced workstation memory usage and improved application startup speed.")

            AppendSeparator()

            AppendHeader("Version 2.1.0 - October 2026")
            AppendBullet("What's New & Release Changelogs: Integrated release history and update notification window.")
            AppendBullet("1-Click System Updates: Self-service update checker querying GitHub Releases with automated in-place restart.")
            AppendBullet("Embedded Assembly Metadata: Official Bangsamoro Transition Authority metadata headers to prevent heuristic antivirus flagging.")
            AppendBullet("Automated Stream Unblocking: Updated binaries automatically strip Windows Mark-of-the-Web zone identifiers.")
            AppendBullet("Audit Trail Hash Chain: Cryptographic SHA-256 seal chains for tamper-evident document action logging.")
            AppendBullet("Multi-Workstation Sync: Background peer heartbeat discovery and non-blocking SQL Server mirror merges.")

            AppendSeparator()

            AppendHeader("Version 2.0.0 - September 2026")
            AppendBullet("External Intake Web Portal: Citizen and ministry intake portal with OTP verification and tracking.")
            AppendBullet("Routing Slip Print Engine: Standard Bangsamoro parliamentary document routing slip print service.")
            AppendBullet("Executive Action Directives: Secretary-General revision loops and status routing.")
            AppendBullet("Controlled Storage Tracking: Cabinet, shelf, and box physical archive movement tracking.")
            AppendBullet("High-DPI Civic Calm Theme: Modern, accessible, high-contrast WinForms visual design.")

            AppendSeparator()

            AppendHeader("Version 1.5.0 - August 2026")
            AppendBullet("RFID Dual-Factor Security: Contactless smart card authentication with terminal auto-lockout defense.")
            AppendBullet("Offline Resilient Outbox: Continuous background sync to SQL Server with seamless offline cache fallback.")
            AppendBullet("Role-Based Access Control: Granular permissions for Records Officers, Legal Counsel, and Office of the SG.")

            AppendSeparator()

            AppendHeader("Version 1.0.0 - July 2026")
            AppendBullet("Document Tracking Engine: Central database tracking for communications, bills, vouchers, and travel orders.")
            AppendBullet("Sequential Tracking Numbers: Automated document tracking code generator (COMM, LEG, FIN, TO).")
            AppendBullet("OSG Desk Routing: Digital document routing and status transitions between division desks.")

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

        Private Sub AppendBody(text As String)
            rtbChangelog.SelectionFont = CivicCalmTheme.FontBody
            rtbChangelog.SelectionColor = CivicCalmTheme.ColorInkMuted
            rtbChangelog.AppendText(text & vbCrLf)
        End Sub

        Private Sub AppendSeparator()
            rtbChangelog.AppendText(vbCrLf)
        End Sub
    End Class
End Namespace
