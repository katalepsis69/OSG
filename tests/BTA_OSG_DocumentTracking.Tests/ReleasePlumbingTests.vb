Option Explicit On
Option Strict On

Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports BTA_OSG

Namespace BTA_OSG.Tests
    ''' <summary>
    ''' Release plumbing in dotnet test, not only in the self-check: stamp parsing,
    ''' same-version republish detection, and two-part version normalization are pure
    ''' functions, so the merge gate can run them directly.
    ''' </summary>
    <TestClass>
    Public Class ReleasePlumbingTests

        <TestMethod>
        Public Sub ParseBuildStamp_ReadsStampFromReleaseBody()
            Dim body = "### What's New in v2.1.6" & vbCrLf & vbCrLf & "- Fixed a thing" & vbCrLf & vbCrLf & "<!-- build:2026-10-03T14:22:11Z -->"
            Assert.AreEqual(New DateTime(2026, 10, 3, 14, 22, 11, DateTimeKind.Utc), AppUpdateService.ParseBuildStamp(body))
        End Sub

        <TestMethod>
        Public Sub ParseBuildStamp_UnstampedBodyYieldsMinValue()
            Assert.AreEqual(DateTime.MinValue, AppUpdateService.ParseBuildStamp("notes without a stamp"))
            Assert.AreEqual(DateTime.MinValue, AppUpdateService.ParseBuildStamp(Nothing))
        End Sub

        <TestMethod>
        Public Sub ParseInformationalStamp_StripsSdkSourceRevisionSuffix()
            Dim expected = New DateTime(2026, 10, 3, 14, 22, 11, DateTimeKind.Utc)
            Assert.AreEqual(expected, AppUpdateService.ParseInformationalStamp("2.1.6+2026-10-03T14:22:11Z.d196c725"))
            Assert.AreEqual(DateTime.MinValue, AppUpdateService.ParseInformationalStamp("2.1.6"))
        End Sub

        <TestMethod>
        Public Sub IsRepublishNewer_OnlyANewerStampCounts()
            Dim newer = New DateTime(2026, 10, 3, 14, 22, 11, DateTimeKind.Utc)
            Dim older = newer.AddMinutes(-1)
            Assert.IsTrue(AppUpdateService.IsRepublishNewer(newer, older))
            Assert.IsFalse(AppUpdateService.IsRepublishNewer(older, newer))
            Assert.IsFalse(AppUpdateService.IsRepublishNewer(DateTime.MinValue, older))
        End Sub

        <TestMethod>
        Public Sub IsSameVersion_NormalizesTwoPartTags()
            Assert.IsTrue(AppUpdateService.IsSameVersion(New Version(2, 2), New Version(2, 2, 0, 0)))
            Assert.IsTrue(AppUpdateService.IsSameVersion(New Version(2, 1, 6), New Version(2, 1, 6, 0)))
            Assert.IsFalse(AppUpdateService.IsSameVersion(New Version(2, 2), New Version(2, 1, 6, 0)))
        End Sub

        <TestMethod>
        Public Sub FormatVersion_AlwaysTwoParts()
            Assert.AreEqual("2.2", AppUpdateService.FormatVersion(New Version(2, 2, 0, 0)))
            Assert.AreEqual("2.1", AppUpdateService.FormatVersion(New Version(2, 1, 6, 0)))
        End Sub

        <TestMethod>
        Public Sub FormatTagLabel_ReducesToTwoParts()
            Assert.AreEqual("v2.1", AppUpdateService.FormatTagLabel("v2.1.5"))
            Assert.AreEqual("v2.2", AppUpdateService.FormatTagLabel("v2.2"))
            Assert.AreEqual("weird-tag", AppUpdateService.FormatTagLabel("weird-tag"))
        End Sub

        <TestMethod>
        Public Sub NormalizeVersion_MatchesTwoPartTagsAgainstCuratedEntries()
            Assert.AreEqual(FormWhatsNew.NormalizeVersion("v2.1"), FormWhatsNew.NormalizeVersion("2.1.0.0"))
            Assert.AreNotEqual(FormWhatsNew.NormalizeVersion("v2.1"), FormWhatsNew.NormalizeVersion("v2.1.6"))
        End Sub
    End Class
End Namespace
