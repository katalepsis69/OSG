Option Explicit On
Option Strict On

Imports System.Drawing
Imports System.Windows.Forms

Namespace BTA_OSG
    ''' Panels and PrintPreviewControl only receive the mouse wheel when they can hold focus,
    ''' and neither can, so they scroll solely by dragging the thumb. This routes the wheel to
    ''' whatever the cursor is actually hovering over.
    Public Class WheelScroller
        Implements IMessageFilter

        Private Const WM_MOUSEWHEEL As Integer = &H20A
        Private Const WheelNotch As Integer = 120

        Public Shared Sub Install()
            Application.AddMessageFilter(New WheelScroller())
        End Sub

        Public Function PreFilterMessage(ByRef m As Message) As Boolean Implements IMessageFilter.PreFilterMessage
            If m.Msg <> WM_MOUSEWHEEL Then Return False

            ' A message filter runs on every wheel event for the life of the app; any throw
            ' here surfaces as an app-wide System Error dialog, so it must fail open.
            Try
                Dim wheelDelta As Integer = WheelDeltaFromWParam(m.WParam)
                Dim down As Boolean = wheelDelta < 0

                Dim hovered = ControlAtCursor()
                If hovered Is Nothing Then Return False

                ' Only controls that scroll their own content may keep the wheel. Bailing out on
                ' every interactive control left the wheel dead wherever the cursor happened to
                ' sit on a text box, combo, or date picker, which is most of a form.
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
                    ' AutoScrollPosition reports a negative offset, so scrolling down subtracts.
                    Dim current = panel.AutoScrollPosition
                    Dim notches As Integer = Math.Max(1, Math.Abs(wheelDelta) \ WheelNotch)
                    Dim proposed As New Point(current.X, current.Y + If(down, -1, 1) * notches * WheelNotch)
                    If proposed.Y = current.Y Then Return False
                    panel.AutoScrollPosition = proposed
                    Return True
                End If

                Return False
            Catch
                Return False
            End Try
        End Function

        ''' The delta is the signed high word of wParam. CShort would throw here on wheel-down
        ''' (0xFF88 does not fit a Short under VB's checked conversion), so sign-extend by hand.
        Public Shared Function WheelDeltaFromWParam(wParam As IntPtr) As Integer
            Dim hiWord As Long = (wParam.ToInt64() >> 16) And &HFFFF
            Return CInt(If(hiWord >= &H8000, hiWord - &H10000, hiWord))
        End Function

        ''' Top-most first: with a modal dialog open, the main form still contains the cursor point
        ''' and GetChildAtPoint ignores overlapping windows.
        Private Shared Function ControlAtCursor() As Control
            Dim screenPoint As Point = Cursor.Position
            Dim forms = Application.OpenForms
            For i As Integer = forms.Count - 1 To 0 Step -1
                Dim openForm = forms(i)
                If openForm.Bounds.Contains(screenPoint) Then
                    Dim found = openForm.GetChildAtPoint(openForm.PointToClient(screenPoint), GetChildAtPointSkip.Invisible)
                    If found IsNot Nothing Then Return found
                End If
            Next
            Return Nothing
        End Function

        ''' A grid, list, tree, or multiline editor scrolls itself; a single-line box, combo,
        ''' date picker, or numeric up-down cannot, so those must fall through to the panel.
        Private Shared Function KeepsWheelForItself(start As Control) As Boolean
            Dim current As Control = start
            While current IsNot Nothing
                Dim editor = TryCast(current, TextBoxBase)
                If editor IsNot Nothing AndAlso editor.Multiline Then Return True
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

        ''' TableLayoutPanel and FlowLayoutPanel are Panels, so the first ancestor of that type
        ''' is usually an inner layout container that does not scroll. Keep climbing.
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
