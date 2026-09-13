Imports System
Imports System.Text
Imports System.Windows.Forms

Namespace BTA_OSG
    Public Class RfidInputListener
        Private _buffer As New StringBuilder()
        Private WithEvents _timer As New Timer()
        Public Event CardScanned(cardId As String)

        Public Sub New(form As Form)
            form.KeyPreview = True
            AddHandler form.KeyPress, AddressOf OnKeyPress
            _timer.Interval = 500
        End Sub

        Private Sub OnKeyPress(sender As Object, e As KeyPressEventArgs)
            _timer.Stop()
            _timer.Start()
            
            If e.KeyChar = Convert.ToChar(Keys.Enter) Then
                If _buffer.Length > 0 Then
                    Dim cardId As String = _buffer.ToString().Trim().ToUpper()
                    RaiseEvent CardScanned(cardId)
                    _buffer.Clear()
                End If
            Else
                _buffer.Append(e.KeyChar)
            End If
        End Sub

        Private Sub _timer_Tick(sender As Object, e As EventArgs) Handles _timer.Tick
            _buffer.Clear()
            _timer.Stop()
        End Sub
    End Class
End Namespace
