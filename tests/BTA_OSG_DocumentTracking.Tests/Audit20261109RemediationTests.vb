Option Explicit On
Option Strict On

Imports System
Imports System.Linq
Imports System.Drawing
Imports System.Threading.Tasks
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports BTA_OSG

' Regression coverage for the 2026-11-09 audit remediation: the updater's failure
' reporting (audit findings 2 and 3) and the shared deadline-tint pairs the quick
' view now consumes (audit finding 5). Everything here runs headless.
<TestClass>
Public Class Audit20261109RemediationTests

    <TestMethod>
    Public Async Function UpdateDownloadFailure_ReportsThroughProgressAndReturnsFalse() As Task
        ' The old failure path called MessageBox.Show(ownerForm, ...) directly on the
        ' worker thread the ConfigureAwait(False) continuations resume on; with a
        ' Nothing owner that threw NullReferenceException out of the Catch, and the
        ' progress line never heard about the failure. The fix reports through the
        ' caller-owned progress line and returns False instead of exiting the process.
        Dim labels As New List(Of String)()
        Dim gate As New Object()
        Dim reporter As New Progress(Of DownloadProgressInfo)(Sub(p)
                                                                  SyncLock gate
                                                                      labels.Add(p.Label)
                                                                  End SyncLock
                                                              End Sub)

        Dim result = Await AppUpdateService.DownloadAndApplyAsync("http://127.0.0.1:9/no-such-update.exe", Nothing, reporter)

        Assert.IsFalse(result, "a failed download must report False, not success")
        ' Progress<T> delivery is asynchronous on purpose; poll briefly for the report.
        Dim deadline = DateTime.UtcNow.AddSeconds(5)
        Do While DateTime.UtcNow < deadline AndAlso Not labels.Any(Function(l) l IsNot Nothing AndAlso l.StartsWith("Update failed", StringComparison.Ordinal))
            Await Task.Delay(50)
        Loop
        Assert.IsTrue(labels.Any(Function(l) l IsNot Nothing AndAlso l.StartsWith("Update failed", StringComparison.Ordinal)),
                      "the failure reaches the visible progress line instead of a cross-thread dialog")
    End Function

    <TestMethod>
    Public Sub DeadlineTint_QuickViewPairs_StayTheSharedDecision()
        ' The quick view delegates to the shared tint and renders its foreground only.
        ' These pins lock the exact pairs the dialog shows, so a future threshold or
        ' pair change inside DeadlineTint cannot silently re-tint the quick view
        ' (the pre-fix quick view carried its own inline copy that could drift).
        Dim bg As Color = Nothing
        Dim fg As Color = Nothing

        Dim overdue = DateTime.Now.AddHours(-30).ToString("yyyy-MM-dd HH:mm:ss", Globalization.CultureInfo.InvariantCulture)
        Assert.IsTrue(DataGridStyler.DeadlineTint(overdue, "FOR_REVIEW", bg, fg))
        Assert.AreEqual(CivicCalmTheme.ColorDanger, fg, "overdue foreground stays the danger ink the quick view renders")

        Dim dueSoon = DateTime.Now.AddHours(4).ToString("yyyy-MM-dd HH:mm:ss", Globalization.CultureInfo.InvariantCulture)
        Assert.IsTrue(DataGridStyler.DeadlineTint(dueSoon, "FOR_REVIEW", bg, fg))
        Assert.AreEqual(CivicCalmTheme.ColorStatusReceivedFg, fg, "due-soon foreground stays the amber ink the quick view renders")
    End Sub

End Class
