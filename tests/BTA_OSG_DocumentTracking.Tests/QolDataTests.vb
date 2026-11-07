Option Explicit On
Option Strict On

Imports System
Imports System.Drawing
Imports System.Linq
Imports System.Text.Json
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports BTA_OSG

' Behavioral coverage for the 2026-11-06 QOL batch's non-UI surfaces (audit finding 8.1):
' the derived filter vocabulary, the two new GetVisibleDocuments parameters, the UiState
' settings round-trip, and the deadline/status tint decisions. Everything here runs
' headless on the offline harness; the paint-time application is inspection-verified.
<TestClass>
Public Class QolDataTests

    <TestMethod>
    Public Sub FilterableStatusCodes_DerivesFromCanonicalLists()
        Dim codes = DocumentStatus.FilterableStatusCodes
        ' RECEIVED (the registration default) plus every routing target plus CANCELLED.
        Assert.IsTrue(codes.Contains("RECEIVED"), "the registration default is filterable")
        Assert.IsTrue(codes.Contains("CANCELLED"), "the voided end state is filterable")
        For Each target As String In DocumentStatus.RouteTargetStatusCodes
            Assert.IsTrue(codes.Contains(target), target & " must ride along from the canonical route list")
        Next
        Assert.AreEqual(DocumentStatus.RouteTargetStatusCodes.Length + 2, codes.Length, "derived, never restated: route targets plus the two bookends")
        Assert.AreEqual(codes.Length, codes.Distinct().Count(), "no duplicate codes")
    End Sub

    <TestMethod>
    Public Sub FilterableStatusCodes_EveryCodeHasDisplayName()
        For Each code As String In DocumentStatus.FilterableStatusCodes
            Assert.IsFalse(String.IsNullOrWhiteSpace(DocumentStatus.DisplayName(code)), code & " must present a display name")
        Next
    End Sub

    <TestMethod>
    Public Sub GetVisibleDocuments_FiltersByAssignedStaff()
        EmbeddedDB.Initialize()
        EmbeddedDB.AddUser("QOLSG00001", "QOL Test Secretary-General", "Secretary-General", "Office of the Secretary-General")
        Dim userSG = EmbeddedDB.AuthenticateRFID("QOLSG00001")
        Assert.IsNotNull(userSG)

        EmbeddedDB.AddDocument("QOLSTAFF01", "Regular Communication", "QOL staff filter mine", "Origin A", "Dest A", "CAB-A", "S-1", "BOX-01", "", "FOR_REVIEW", "QOL Test Secretary-General", "INCOMING", "Secretariat")
        EmbeddedDB.AddDocument("QOLSTAFF02", "Regular Communication", "QOL staff filter other", "Origin B", "Dest B", "CAB-A", "S-1", "BOX-01", "", "FOR_REVIEW", "QOL Other Officer", "INCOMING", "Secretariat")

        Dim mine = EmbeddedDB.GetVisibleDocuments(userSG, "", "", "QOL Test Secretary-General", "")
        Dim mineTitles = mine.Cast(Of System.Data.DataRowView)().Select(Function(r) r("Title").ToString()).ToList()
        Assert.IsTrue(mineTitles.Contains("QOL staff filter mine"), "the staff filter keeps the officer's own row")
        Assert.IsFalse(mineTitles.Contains("QOL staff filter other"), "the staff filter drops another officer's row")
    End Sub

    <TestMethod>
    Public Sub GetVisibleDocuments_FiltersByStatusCode()
        EmbeddedDB.Initialize()
        EmbeddedDB.AddUser("QOLSG00002", "QOL Test Secretary-General Two", "Secretary-General", "Office of the Secretary-General")
        Dim userSG = EmbeddedDB.AuthenticateRFID("QOLSG00002")
        Assert.IsNotNull(userSG)

        EmbeddedDB.AddDocument("QOLSTAT01", "Regular Communication", "QOL status probe review", "Origin C", "Dest C", "CAB-B", "S-2", "BOX-02", "", "FOR_REVIEW", "QOL Test Secretary-General Two", "INCOMING", "Secretariat")
        EmbeddedDB.AddDocument("QOLSTAT02", "Finance", "QOL status probe received", "Origin D", "Dest D", "CAB-B", "S-2", "BOX-02", "", "RECEIVED", "QOL Test Secretary-General Two", "INCOMING", "Finance Section")

        Dim forReview = EmbeddedDB.GetVisibleDocuments(userSG, "", "", "", "FOR_REVIEW")
        Dim titles = forReview.Cast(Of System.Data.DataRowView)().Select(Function(r) r("Title").ToString()).ToList()
        Assert.IsTrue(titles.Contains("QOL status probe review"), "the status filter keeps the matching row")
        Assert.IsFalse(titles.Contains("QOL status probe received"), "the status filter drops other statuses")

        ' Staff and status compose.
        Dim combined = EmbeddedDB.GetVisibleDocuments(userSG, "", "", "QOL Test Secretary-General Two", "RECEIVED")
        Dim combinedTitles = combined.Cast(Of System.Data.DataRowView)().Select(Function(r) r("Title").ToString()).ToList()
        Assert.IsTrue(combinedTitles.Contains("QOL status probe received"), "both filters together keep the row matching both")
        Assert.IsFalse(combinedTitles.Contains("QOL status probe review"), "both filters together drop the status mismatch")
    End Sub

    <TestMethod>
    Public Sub UiStateSettings_JsonRoundTrip()
        Dim ui As New UiStateSettings With {
            .DashboardSection = "Records Section",
            .DashboardCategory = "Finance",
            .DashboardStatus = "For Review",
            .DashboardOnlyMine = True
        }
        Dim root As New AppSettingsRoot With {.UiState = ui}
        Dim options As New JsonSerializerOptions With {.WriteIndented = True}
        Dim json = JsonSerializer.Serialize(root, options)

        Dim back = JsonSerializer.Deserialize(Of AppSettingsRoot)(json, options)
        Assert.IsNotNull(back)
        Assert.IsNotNull(back.UiState, "the UiState layer survives the round trip")
        Assert.AreEqual("Records Section", back.UiState.DashboardSection)
        Assert.AreEqual("Finance", back.UiState.DashboardCategory)
        Assert.AreEqual("For Review", back.UiState.DashboardStatus)
        Assert.IsTrue(back.UiState.DashboardOnlyMine)
    End Sub

    <TestMethod>
    Public Sub UiStateSettings_FailsOpenOnEmptyDefaults()
        ' The restore path only applies entries present in the live lists; the shape that
        ' reaches it when nothing was ever saved must be all-default, not null.
        Dim ui As New UiStateSettings()
        Assert.AreEqual("", ui.DashboardSection)
        Assert.AreEqual("", ui.DashboardCategory)
        Assert.AreEqual("", ui.DashboardStatus)
        Assert.IsFalse(ui.DashboardOnlyMine)
    End Sub

    <TestMethod>
    Public Sub StatusColors_KnownCodesMapAndUnknownRefuses()
        Dim bg As Color = Nothing
        Dim fg As Color = Nothing

        Assert.IsTrue(DataGridStyler.StatusColors("FOR_REVIEW", bg, fg))
        Assert.AreEqual(CivicCalmTheme.ColorStatusReviewBg, bg)

        Assert.IsTrue(DataGridStyler.StatusColors("for_revision", bg, fg), "the lookup is case-insensitive")
        Assert.AreEqual(CivicCalmTheme.ColorStatusRevisionFg, fg)

        Assert.IsTrue(DataGridStyler.StatusColors("RECEIVED", bg, fg))
        Assert.AreEqual(CivicCalmTheme.ColorStatusReceivedBg, bg)

        Assert.IsFalse(DataGridStyler.StatusColors("TOTALLY_UNKNOWN", bg, fg), "unknown codes refuse so callers keep their stored text")
        Assert.IsFalse(DataGridStyler.StatusColors("", bg, fg), "empty refuses")
    End Sub

    <TestMethod>
    Public Sub DeadlineTint_OverdueDueSoonAndTerminal()
        Dim bg As Color = Nothing
        Dim fg As Color = Nothing

        Dim overdue = DateTime.Now.AddHours(-30).ToString("yyyy-MM-dd HH:mm:ss", Globalization.CultureInfo.InvariantCulture)
        Assert.IsTrue(DataGridStyler.DeadlineTint(overdue, "FOR_REVIEW", bg, fg), "an overdue live document tints")
        Assert.AreEqual(CivicCalmTheme.ColorDangerSoft, bg)
        Assert.AreEqual(CivicCalmTheme.ColorDanger, fg)

        Dim dueSoon = DateTime.Now.AddHours(4).ToString("yyyy-MM-dd HH:mm:ss", Globalization.CultureInfo.InvariantCulture)
        Assert.IsTrue(DataGridStyler.DeadlineTint(dueSoon, "IN_PROGRESS", bg, fg), "due within 24 hours tints amber")
        Assert.AreEqual(CivicCalmTheme.ColorStatusReceivedBg, bg)

        Dim later = DateTime.Now.AddDays(10).ToString("yyyy-MM-dd HH:mm:ss", Globalization.CultureInfo.InvariantCulture)
        Assert.IsFalse(DataGridStyler.DeadlineTint(later, "ROUTED", bg, fg), "comfortable deadlines stay plain")

        Assert.IsFalse(DataGridStyler.DeadlineTint(overdue, "RELEASED", bg, fg), "terminal rows never tint")
        Assert.IsFalse(DataGridStyler.DeadlineTint(overdue, "CANCELLED", bg, fg), "voided rows never tint")
        Assert.IsFalse(DataGridStyler.DeadlineTint("", "FOR_REVIEW", bg, fg), "empty deadline refuses")
        Assert.IsFalse(DataGridStyler.DeadlineTint("not a date", "FOR_REVIEW", bg, fg), "unparseable text refuses")
    End Sub

End Class
