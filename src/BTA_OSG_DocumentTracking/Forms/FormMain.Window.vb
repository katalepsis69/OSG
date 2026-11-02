Option Explicit On
Option Strict On

Imports System
Imports System.Drawing
Imports System.Windows.Forms

Namespace BTA_OSG
    Partial Public Class FormMain

        ' Restore the last placement only when it still lands on a monitor that exists now.
        ' A saved rectangle from a disconnected second screen would otherwise open the app
        ' off-screen, where the operator cannot drag it back.
        Private Sub ApplySavedWindowBounds()
            Dim savedSettings = AppSettings.Instance.WindowSettings
            If savedSettings Is Nothing OrElse Not savedSettings.HasBounds OrElse
               savedSettings.Width <= 0 OrElse savedSettings.Height <= 0 Then Return

            Dim saved As New Rectangle(savedSettings.Left, savedSettings.Top, savedSettings.Width, savedSettings.Height)
            Dim targetScreen As Screen = Nothing
            For Each candidate In Screen.AllScreens
                If candidate.WorkingArea.IntersectsWith(saved) Then
                    targetScreen = candidate
                    Exit For
                End If
            Next
            If targetScreen Is Nothing Then Return

            Dim area = targetScreen.WorkingArea
            Dim width = Math.Min(saved.Width, area.Width)
            Dim height = Math.Min(saved.Height, area.Height)
            Dim left = Math.Max(area.Left, Math.Min(saved.Left, area.Right - width))
            Dim top = Math.Max(area.Top, Math.Min(saved.Top, area.Bottom - height))

            Me.StartPosition = FormStartPosition.Manual
            Me.Bounds = New Rectangle(left, top, width, height)
            If savedSettings.Maximized Then Me.WindowState = FormWindowState.Maximized
        End Sub

        Private Sub SaveWindowBounds()
            Try
                ' RestoreBounds is the normal-state rectangle even when the window is
                ' maximized or minimized at close; Bounds would save the wrong size.
                Dim bounds = If(Me.WindowState = FormWindowState.Normal, Me.Bounds, Me.RestoreBounds)
                If bounds.Width < 200 OrElse bounds.Height < 200 Then Return
                Dim savedSettings = AppSettings.Instance.WindowSettings
                If savedSettings Is Nothing Then
                    savedSettings = New WindowSettings()
                    AppSettings.Instance.WindowSettings = savedSettings
                End If
                savedSettings.HasBounds = True
                savedSettings.Left = bounds.Left
                savedSettings.Top = bounds.Top
                savedSettings.Width = bounds.Width
                savedSettings.Height = bounds.Height
                savedSettings.Maximized = (Me.WindowState = FormWindowState.Maximized)
                AppSettings.Save(AppSettings.Instance)
            Catch ex As Exception
                System.Diagnostics.Trace.TraceWarning("Window bounds could not be saved: " & ex.Message)
            End Try
        End Sub

    End Class
End Namespace
