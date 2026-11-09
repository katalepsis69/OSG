Option Explicit On
Option Strict On

Imports System.Drawing
Imports System.Windows.Forms

Namespace BTA_OSG
    ' Panels and PrintPreviewControl do not receive mouse wheel events without focus;
    ' this message filter routes the wheel to the scrollable container under the cursor.
    Public Class WheelScroller
        Implements IMessageFilter

        Private Const WM_MOUSEWHEEL As Integer = &H20A
        Private Const WheelNotch As Integer = 120

        Public Shared Sub Install()
            Application.AddMessageFilter(New WheelScroller())
        End Sub

        Public Function PreFilterMessage(ByRef m As Message) As Boolean Implements IMessageFilter.PreFilterMessage
            If m.Msg <> WM_MOUSEWHEEL Then Return False

            ' Runs on every wheel event app-wide; fail open rather than raising an unhandled dialog.
            Try
                Dim wheelDelta As Integer = WheelDeltaFromWParam(m.WParam)
                Dim down As Boolean = wheelDelta < 0

                Dim hovered = ControlAtCursor()
                If hovered Is Nothing Then Return False

                If KeepsWheelForItself(hovered) Then Return False

                Dim preview = FindAncestor(Of PrintPreviewControl)(hovered)
                If preview IsNot Nothing Then
                    Dim target As Integer = preview.StartPage + If(down, 1, -1)
                    If target < 0 Then Return False
                    preview.StartPage = target
                    Return True
                End If

                Dim panel = FindScrollableAncestor(hovered)
                If panel IsNot Nothing Then
                    ' AutoScrollPosition returns negative offsets on read but takes positive offsets on write.
                    Dim current = panel.AutoScrollPosition
                    Dim notches As Integer = Math.Max(1, Math.Abs(wheelDelta) \ WheelNotch)
                    Dim proposedTop As Integer = Math.Max(0, -current.Y + If(down, 1, -1) * notches * WheelNotch)
                    If proposedTop = -current.Y Then Return False
                    panel.AutoScrollPosition = New Point(-current.X, proposedTop)
                    Return True
                End If

                Return False
            Catch
                Return False
            End Try
        End Function

        ' Signed high word of wParam; sign-extended manually because 0xFF88 overflows checked Short.
        Public Shared Function WheelDeltaFromWParam(wParam As IntPtr) As Integer
            Dim hiWord As Long = (wParam.ToInt64() >> 16) And &HFFFF
            Return CInt(If(hiWord >= &H8000, hiWord - &H10000, hiWord))
        End Function

        ' Top-most form first, then drill down to the leaf child under the cursor.
        Private Shared Function ControlAtCursor() As Control
            Dim screenPoint As Point = Cursor.Position
            Dim forms = Application.OpenForms
            For i As Integer = forms.Count - 1 To 0 Step -1
                Dim openForm = forms(i)
                If openForm.Bounds.Contains(screenPoint) Then
                    Dim current As Control = openForm
                    While True
                        Dim child = current.GetChildAtPoint(current.PointToClient(screenPoint), GetChildAtPointSkip.Invisible)
                        If child Is Nothing OrElse child Is current Then Exit While
                        current = child
                    End While
                    If current IsNot openForm Then Return current
                End If
            Next
            Return Nothing
        End Function

        ' Scrollable grids, lists, trees, and editors with scrollbars keep the wheel; non-scrolling controls fall through.
        Private Shared Function KeepsWheelForItself(start As Control) As Boolean
            Dim current As Control = start
            While current IsNot Nothing
                Dim rtb = TryCast(current, RichTextBox)
                If rtb IsNot Nothing Then
                    If rtb.ScrollBars <> RichTextBoxScrollBars.None Then Return True
                Else
                    Dim editor = TryCast(current, TextBoxBase)
                    If editor IsNot Nothing AndAlso editor.Multiline Then Return True
                End If
                If TypeOf current Is DataGridView OrElse TypeOf current Is ListBox OrElse TypeOf current Is TreeView Then Return True
                current = current.Parent
            End While
            Return False
        End Function

        Private Shared Function FindAncestor(Of T As Control)(start As Control) As T
            Dim current As Control = start
            While current IsNot Nothing
                If TypeOf current Is T Then Return TryCast(current, T)
                current = current.Parent
            End While
            Return Nothing
        End Function

        Private Shared Function FindScrollableAncestor(start As Control) As Panel
            Dim current As Control = start
            While current IsNot Nothing
                Dim candidate = TryCast(current, Panel)
                If candidate IsNot Nothing AndAlso candidate.AutoScroll Then Return candidate
                current = current.Parent
            End While
            Return Nothing
        End Function
    End Class
End Namespace
