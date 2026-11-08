Option Explicit On
Option Strict On

Imports System
Imports System.Diagnostics
Imports System.Drawing
Imports System.Windows.Forms
Imports QRCoder

Namespace BTA_OSG
    ''' <summary>
    ''' Civic Calm dialog for the Cloudflare tunnel. Renders the public URL as a
    ''' scannable QR code beside the link text, so the operator's phone becomes the
    ''' reader instead of the keyboard. Replaces the plain MessageBox the tunnel
    ''' actions used to raise (AGENTS Rule 7: no decorative marks, keyboard and
    ''' DPI-safe layout via docking containers).
    ''' </summary>
    Public Class FormTunnelStatus
        Inherits Form

        ''' <summary>JustStarted: the tunnel came up during this click. AlreadyRunning:
        ''' an active tunnel was clicked again, so the operator can stop it here.</summary>
        Public Enum TunnelContext
            JustStarted
            AlreadyRunning
        End Enum

        Private ReadOnly _publicUrl As String
        Private ReadOnly _context As TunnelContext

        Private pctQr As PictureBox
        Private txtUrl As TextBox
        Private lblNote As Label
        Private flpFooter As FlowLayoutPanel
        Private btnPrimary As Button
        Private btnSecondary As Button

        Public Sub New(publicUrl As String, context As TunnelContext)
            If String.IsNullOrWhiteSpace(publicUrl) Then Throw New ArgumentException("Tunnel URL is required.", NameOf(publicUrl))
            _publicUrl = publicUrl
            _context = context
            InitializeComponent()
        End Sub

        Private Sub InitializeComponent()
            Me.Text = If(_context = TunnelContext.AlreadyRunning,
                         "Cloudflare Tunnel Running - BTA OSG Document Tracking",
                         "Cloudflare Tunnel Online - BTA OSG Document Tracking")
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.ShowInTaskbar = False
            Me.StartPosition = FormStartPosition.CenterParent
            Me.BackColor = CivicCalmTheme.ColorSurface
            Me.Font = CivicCalmTheme.FontBody
            Me.Size = New Size(380, 540)
            Me.MinimumSize = New Size(360, 500)

            ' Body: QR plate, URL box, microcopy
            Dim pnlBody As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.TopDown,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(18, 16, 18, 8),
                .WrapContents = False,
                .AutoScroll = False
            }

            pctQr = New PictureBox With {
                .SizeMode = PictureBoxSizeMode.Normal,
                .AutoSize = True,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Margin = New Padding(56, 0, 56, 14)
            }
            pnlBody.Controls.Add(pctQr)

            Dim lblUrlCaption As New Label With {
                .Text = "Public URL:",
                .Font = CivicCalmTheme.FontFieldLabel,
                .ForeColor = CivicCalmTheme.ColorInk,
                .AutoSize = True,
                .Margin = New Padding(0, 0, 0, 4)
            }
            pnlBody.Controls.Add(lblUrlCaption)

            txtUrl = New TextBox With {
                .Text = _publicUrl,
                .ReadOnly = True,
                .Width = 320,
                .Font = CivicCalmTheme.FontIdentifier,
                .BackColor = CivicCalmTheme.ColorWell,
                .BorderStyle = BorderStyle.FixedSingle,
                .Margin = New Padding(0, 0, 0, 8)
            }
            pnlBody.Controls.Add(txtUrl)

            lblNote = New Label With {
                .Text = "(The link has been copied to your clipboard.)",
                .Font = CivicCalmTheme.FontMicrocopy,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .AutoSize = True,
                .Margin = New Padding(0, 0, 0, 0)
            }
            pnlBody.Controls.Add(lblNote)

            ' Footer buttons
            Dim pnlFooter As New Panel With {
                .Dock = DockStyle.Bottom,
                .Height = 62,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(16, 12, 16, 12)
            }
            flpFooter = New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.RightToLeft,
                .BackColor = CivicCalmTheme.ColorSurface
            }

            If _context = TunnelContext.AlreadyRunning Then
                btnSecondary = CivicCalmTheme.CreateIconButton("&Keep Running", margin:=New Padding(6, 0, 0, 0))
                AddHandler btnSecondary.Click, Sub()
                                                   Me.DialogResult = DialogResult.No
                                                   Me.Close()
                                               End Sub
                flpFooter.Controls.Add(btnSecondary)

                ' The factory supplies the white foreground for a filled background, so the
                ' danger button keeps its contrast without a hand-built Color.White.
                btnPrimary = CivicCalmTheme.CreateIconButton("&Stop Tunnel", backColor:=CivicCalmTheme.ColorDanger)
                AddHandler btnPrimary.Click, Sub()
                                                 Me.DialogResult = DialogResult.Yes
                                                 Me.Close()
                                             End Sub
                flpFooter.Controls.Add(btnPrimary)
                Me.AcceptButton = btnSecondary
                Me.CancelButton = btnSecondary
            Else
                btnSecondary = CivicCalmTheme.CreateIconButton("&Close", margin:=New Padding(6, 0, 0, 0))
                AddHandler btnSecondary.Click, Sub()
                                                   Me.DialogResult = DialogResult.Cancel
                                                   Me.Close()
                                               End Sub
                flpFooter.Controls.Add(btnSecondary)

                btnPrimary = CivicCalmTheme.CreateIconButton("&Open in Browser", backColor:=CivicCalmTheme.ColorPrimary)
                AddHandler btnPrimary.Click, Sub() OpenPublicUrl()
                flpFooter.Controls.Add(btnPrimary)
                Me.AcceptButton = btnSecondary
                Me.CancelButton = btnSecondary
            End If

            pnlFooter.Controls.Add(flpFooter)

            Me.Controls.Add(pnlBody)
            Me.Controls.Add(pnlFooter)

            RenderQrCode()
        End Sub

        ''' <summary>Draws the QR at a fixed module scale; a failure must not take the
        ''' dialog down, the URL box above remains the fallback (C-4 resilience).</summary>
        Private Sub RenderQrCode()
            Try
                Using generator As New QRCodeGenerator()
                    Dim data = generator.CreateQrCode(_publicUrl, QRCodeGenerator.ECCLevel.M)
                    Using qr As New QRCode(data)
                        ' Dark modules carry the theme ink so a token change cannot silently
                        ' desync the QR; the white quiet zone stays literal (scanner contrast).
                        Dim inkHex = Drawing.ColorTranslator.ToHtml(CivicCalmTheme.ColorInk)
                        ' ~50-character URL lands at version 3-4; 5px per module plus the
                        ' quiet zone puts the bitmap near 200px, readable by any phone camera.
                        pctQr.Image = qr.GetGraphic(5, inkHex, "#FFFFFF")
                    End Using
                End Using
            Catch
                pctQr.Visible = False
                lblNote.Text = "(QR rendering unavailable; the link above is selectable and already copied.)"
            End Try
        End Sub

        Private Sub OpenPublicUrl()
            Try
                Process.Start(New ProcessStartInfo(_publicUrl) With {.UseShellExecute = True})
                Me.DialogResult = DialogResult.OK
                Me.Close()
            Catch ex As Exception
                MessageBox.Show("Could not open the browser: " & ex.Message, "Cloudflare Tunnel",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Try
        End Sub
    End Class
End Namespace
