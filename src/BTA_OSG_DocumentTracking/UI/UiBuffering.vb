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
    End Class
End Namespace
