Option Explicit On
Option Strict On

Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports BTA_OSG

Namespace BTA_OSG.Tests
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

        <TestMethod>
        Public Sub AppendMarkdownBody_JoinsWrappedBulletsAndPreservesSubheaders()
            Dim md = "### What's New in v3.2" & vbLf & vbLf &
                     "#### Keyboard" & vbLf &
                     "- Alt+C closes the What's New window again. Download stop button renamed Stop Download," & vbLf &
                     "  shortcut Alt+S (old Cancel button fought Close for the same shortcut)." & vbLf &
                     "- Enter and Escape close the What's New window." & vbLf & vbLf &
                     "#### Fixes" & vbLf &
                     "- Closing the What's New window during an update download now stops the download." & vbLf & vbLf &
                     "<!-- build:2026-11-08T23:01:39Z -->"

            Using rtb As New RichTextBox()
                Dim h = rtb.Handle
                FormWhatsNew.AppendMarkdownBody(rtb, md)
                Dim rendered = rtb.Text

                Assert.IsFalse(rendered.Contains("What's New in v3.2"), "Top release heading is already shown in the section header")
                Assert.IsFalse(rendered.Contains("build:"), "Build stamp comment must not appear in rendered text")
                Assert.IsTrue(rendered.Contains("Keyboard"), "Category subheader Keyboard must be rendered")
                Assert.IsTrue(rendered.Contains("Fixes"), "Category subheader Fixes must be rendered")
                Assert.IsTrue(rendered.Contains("Stop Download, shortcut Alt+S"), "Hard-wrapped bullet lines must join into one paragraph")

                ' Subheader after bullets must not inherit the bullet's hanging indent.
                Dim fixesIdx = rendered.IndexOf("Fixes", StringComparison.Ordinal)
                rtb.Select(fixesIdx, 5)
                Assert.AreEqual(0, rtb.SelectionHangingIndent, "Subheader must not inherit bullet hanging indent")
                Assert.IsTrue(rtb.SelectionFont.Bold, "Subheader must be bold")
            End Using
        End Sub
    End Class
End Namespace
