Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass>
Public Class WheelScrollerTests
    <TestMethod>
    Public Sub WheelDelta_ScrollUp_IsPositiveNotch()
        Assert.AreEqual(120, BTA_OSG.WheelScroller.WheelDeltaFromWParam(New IntPtr(&H780000L)))
    End Sub

    <TestMethod>
    Public Sub WheelDelta_ScrollDown_IsNegativeNotch()
        Assert.AreEqual(-120, BTA_OSG.WheelScroller.WheelDeltaFromWParam(New IntPtr(&HFF880000L)))
    End Sub

    <TestMethod>
    Public Sub WheelDelta_FastScrollDown_IsMultipleNotches()
        Assert.AreEqual(-240, BTA_OSG.WheelScroller.WheelDeltaFromWParam(New IntPtr(&HFF100000L)))
    End Sub

    <TestMethod>
    Public Sub WheelDelta_SignExtendedPointer_MatchesZeroExtended()
        ' Some pumps hand the wParam sign-extended to 64 bits; both forms must agree.
        Dim zeroExtended As New IntPtr(&HFF880000L)
        Dim signExtended As New IntPtr(&HFFFFFFFFFF880000L)
        Assert.AreEqual(BTA_OSG.WheelScroller.WheelDeltaFromWParam(zeroExtended),
                        BTA_OSG.WheelScroller.WheelDeltaFromWParam(signExtended))
    End Sub
End Class
