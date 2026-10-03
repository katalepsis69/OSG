Option Explicit On
Option Strict On

Imports System.Windows.Forms

Namespace BTA_OSG
    ''' <summary>
    ''' Tracks real user input (mouse and keyboard) so the session timeout can measure
    ''' inactivity instead of guessing. Installed as an application message filter: it runs
    ''' on the UI thread, never eats a message, and costs one timestamp write per input
    ''' message.
    ''' </summary>
    Public Class UiActivityMonitor
        Implements IMessageFilter

        Private _lastInputUtc As DateTime = DateTime.UtcNow

        Public Shared ReadOnly Instance As New UiActivityMonitor()

        Public Shared Sub Install()
            Application.AddMessageFilter(Instance)
        End Sub

        Public Function PreFilterMessage(ByRef m As Message) As Boolean Implements IMessageFilter.PreFilterMessage
            Select Case m.Msg
                Case &H100, &H101, &H104, &H105, &H200, &H201, &H204, &H20A
                    ' WM_KEYDOWN, WM_KEYUP, WM_SYSKEYDOWN, WM_SYSKEYUP,
                    ' WM_MOUSEMOVE, WM_LBUTTONDOWN, WM_RBUTTONDOWN, WM_MOUSEWHEEL
                    _lastInputUtc = DateTime.UtcNow
            End Select
            Return False
        End Function

        Public ReadOnly Property IdleSeconds As Integer
            Get
                Return CInt((DateTime.UtcNow - _lastInputUtc).TotalSeconds)
            End Get
        End Property
    End Class
End Namespace
