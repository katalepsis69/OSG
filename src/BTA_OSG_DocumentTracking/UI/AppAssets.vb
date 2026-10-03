Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms

Namespace BTA_OSG
    Public NotInheritable Class AppAssets
        Private Sub New()
        End Sub

        Private Shared _appIcon As Icon = Nothing
        Private Shared _logoImage As Image = Nothing
        Private Shared ReadOnly _crispCache As New Dictionary(Of Integer, Image)()
        Private Shared ReadOnly _syncLock As New Object()

        Public Shared ReadOnly Property AppIcon As Icon
            Get
                SyncLock _syncLock
                    If _appIcon Is Nothing Then
                        Try
                            Dim baseDir = AppDomain.CurrentDomain.BaseDirectory
                            Dim iconPath = Path.Combine(baseDir, "Resources", "app.ico")
                            If File.Exists(iconPath) Then
                                _appIcon = New Icon(iconPath)
                            Else
                                Dim proc = Diagnostics.Process.GetCurrentProcess()
                                If proc.MainModule IsNot Nothing Then
                                    _appIcon = Icon.ExtractAssociatedIcon(proc.MainModule.FileName)
                                End If
                            End If
                        Catch
                            _appIcon = Nothing
                        End Try
                    End If
                    Return _appIcon
                End SyncLock
            End Get
        End Property

        Public Shared ReadOnly Property Logo As Image
            Get
                SyncLock _syncLock
                    If _logoImage Is Nothing Then
                        Try
                            Dim baseDir = AppDomain.CurrentDomain.BaseDirectory
                            Dim logoPath = Path.Combine(baseDir, "Resources", "bta_logo.png")
                            If File.Exists(logoPath) Then
                                Using fs As New FileStream(logoPath, FileMode.Open, FileAccess.Read, FileShare.Read)
                                    Using img = Image.FromStream(fs)
                                        _logoImage = New Bitmap(img)
                                    End Using
                                End Using
                            ElseIf AppIcon IsNot Nothing Then
                                _logoImage = AppIcon.ToBitmap()
                            End If
                        Catch
                            _logoImage = Nothing
                        End Try
                    End If
                    Return _logoImage
                End SyncLock
            End Get
        End Property

        Public Shared Function GetCrispLogo(targetSize As Integer) As Image
            SyncLock _syncLock
                If targetSize <= 0 Then targetSize = 64
                If _crispCache.ContainsKey(targetSize) Then
                    Return _crispCache(targetSize)
                End If

                Dim baseDir = AppDomain.CurrentDomain.BaseDirectory

                ' Check for pre-baked high-fidelity sharpened asset
                Dim candidateNames = New String() {
                    $"logo_{targetSize}.png",
                    $"bta_logo_{targetSize}.png"
                }
                For Each candidate In candidateNames
                    Dim candidatePath = Path.Combine(baseDir, "Resources", candidate)
                    If File.Exists(candidatePath) Then
                        Try
                            Using fs As New FileStream(candidatePath, FileMode.Open, FileAccess.Read, FileShare.Read)
                                Using img = Image.FromStream(fs)
                                    Dim loaded As New Bitmap(img)
                                    _crispCache(targetSize) = loaded
                                    Return loaded
                                End Using
                            End Using
                        Catch
                        End Try
                    End If
                Next

                Dim baseImg = Logo
                If baseImg Is Nothing Then Return Nothing

                ' Progressive multi-stage downsample to eliminate aliasing and blur
                Dim cur As Image = baseImg
                Dim curSize = baseImg.Width

                While curSize > (targetSize * 2)
                    Dim nextSize = curSize \ 2
                    Dim stepBmp As New Bitmap(nextSize, nextSize, Imaging.PixelFormat.Format32bppArgb)
                    Using g As Graphics = Graphics.FromImage(stepBmp)
                        g.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
                        g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                        g.PixelOffsetMode = Drawing2D.PixelOffsetMode.HighQuality
                        g.CompositingQuality = Drawing2D.CompositingQuality.HighQuality
                        g.DrawImage(cur, 0, 0, nextSize, nextSize)
                    End Using
                    If cur IsNot baseImg Then cur.Dispose()
                    cur = stepBmp
                    curSize = nextSize
                End While

                Dim finalBmp As New Bitmap(targetSize, targetSize, Imaging.PixelFormat.Format32bppArgb)
                Using fg As Graphics = Graphics.FromImage(finalBmp)
                    fg.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
                    fg.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                    fg.PixelOffsetMode = Drawing2D.PixelOffsetMode.HighQuality
                    fg.CompositingQuality = Drawing2D.CompositingQuality.HighQuality
                    fg.DrawImage(cur, 0, 0, targetSize, targetSize)
                End Using
                If cur IsNot baseImg Then cur.Dispose()

                _crispCache(targetSize) = finalBmp
                Return finalBmp
            End SyncLock
        End Function

        Public Shared Sub ApplyFormIcon(f As Form)
            If f Is Nothing Then Return
            Try
                Dim ico = AppIcon
                If ico IsNot Nothing Then
                    f.Icon = ico
                End If
            Catch
            End Try
        End Sub

        Public Shared Function CreateLogoPictureBox(Optional sizePx As Integer = 64) As Control
            Dim box As New CrispLogoBox With {
                .Size = New Size(sizePx, sizePx),
                .Margin = New Padding(0, 0, 10, 0)
            }
            Return box
        End Function
    End Class

    Public Class CrispLogoBox
        Inherits Control

        Public Sub New()
            SetStyle(ControlStyles.SupportsTransparentBackColor Or
                     ControlStyles.OptimizedDoubleBuffer Or
                     ControlStyles.AllPaintingInWmPaint Or
                     ControlStyles.UserPaint, True)
            Me.BackColor = Color.Transparent
        End Sub

        Protected Overrides Sub OnPaint(e As PaintEventArgs)
            MyBase.OnPaint(e)
            Dim side = Math.Min(Me.ClientSize.Width, Me.ClientSize.Height)
            If side <= 0 Then Return

            Dim g = e.Graphics
            g.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
            g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
            g.PixelOffsetMode = Drawing2D.PixelOffsetMode.HighQuality
            g.CompositingQuality = Drawing2D.CompositingQuality.HighQuality

            Dim img = AppAssets.GetCrispLogo(side)
            If img IsNot Nothing Then
                Dim x = (Me.ClientSize.Width - side) \ 2
                Dim y = (Me.ClientSize.Height - side) \ 2
                Dim destRect As New Rectangle(x, y, side, side)
                g.DrawImage(img, destRect)
            End If
        End Sub
    End Class
End Namespace
