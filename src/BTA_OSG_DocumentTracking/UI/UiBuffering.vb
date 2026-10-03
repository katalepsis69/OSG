Option Explicit On
Option Strict On

Imports System.Reflection
Imports System.Windows.Forms

Namespace BTA_OSG
    ' WinForms repaints every container HWND separately on maximize/restore, so an unbuffered
    ' panel tree flashes through half-painted states (backgrounds and borders lag behind text).
    ' Panels (including TableLayoutPanel, FlowLayoutPanel, and TabPage, which all derive from
    ' Panel) do not expose DoubleBuffered publicly, so it is enabled through the same reflection
    ' pattern the grid styler uses, applied recursively once after a form builds its tree.
    Public NotInheritable Class UiBuffering
        Private Sub New()
        End Sub

        Private Shared ReadOnly DoubleBufferedProp As PropertyInfo =
            GetType(Control).GetProperty("DoubleBuffered", BindingFlags.Instance Or BindingFlags.NonPublic)

        ' WS_EX_COMPOSITED. Applied to a top-level form, Windows renders the entire child
        ' control tree into an offscreen buffer and presents it as one frame, so restore and
        ' maximize repaints cannot leave ghost copies of controls at their old positions.
        ' Panel-level double buffering alone cannot do this because every child HWND (buttons,
        ' labels, tab control) still paints itself directly to the screen in a separate pass.
        Public Const WsExComposited As Integer = &H2000000

        Public Shared Sub EnableDeep(root As Control)
            If root Is Nothing OrElse DoubleBufferedProp Is Nothing Then Return
            If TypeOf root Is Panel Then
                DoubleBufferedProp.SetValue(root, True, Nothing)
            End If
            For Each child As Control In root.Controls
                EnableDeep(child)
            Next
        End Sub

        ' GWL_EXSTYLE is a 32-bit value, so user32's non-Ptr exports are correct on x64 too.
        ' LibraryImport is unusable here: VB partial methods must be Subs, so these two
        ' integer APIs keep the plain DllImport form.
        <System.Runtime.InteropServices.DllImport("user32.dll")>
        Private Shared Function GetWindowLong(hWnd As IntPtr, nIndex As Integer) As Integer
        End Function

        <System.Runtime.InteropServices.DllImport("user32.dll")>
        Private Shared Function SetWindowLong(hWnd As IntPtr, nIndex As Integer, dwNewLong As Integer) As Integer
        End Function

        Private Const GwlExStyle As Integer = -20

        ' Toggles WS_EX_COMPOSITED on a live handle; CreateParams can only set the style at
        ' creation, but dialogs need it conditionally. Composited repaints make minimize/
        ' maximize/restore atomic, yet while an embedded WebView2 is on screen the same
        ' style re-presents the whole tree on every web frame, which reads as blinking.
        Public Shared Sub SetComposited(form As Form, composited As Boolean)
            If form Is Nothing OrElse Not form.IsHandleCreated Then Return
            Dim exStyle = GetWindowLong(form.Handle, GwlExStyle)
            Dim updated = If(composited, exStyle Or WsExComposited, exStyle And (Not WsExComposited))
            If updated <> exStyle Then
                SetWindowLong(form.Handle, GwlExStyle, updated)
            End If
        End Sub
    End Class
End Namespace
