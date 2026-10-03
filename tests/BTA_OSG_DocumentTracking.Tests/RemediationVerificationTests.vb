Option Explicit On
Option Strict On

Imports System
Imports System.IO
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports BTA_OSG

''' <summary>
''' Assembly-level isolation: the offline XML store is a static, process-wide singleton,
''' and the audit-009 remediation made Initialize actually RELOAD it from disk. Redirect
''' the store at a scratch file so every test class starts from pristine seeds and no
''' stale bin-folder cache can flip an assertion between runs.
''' </summary>
<TestClass>
Public Class RemediationAssemblyInit
    Private Shared _scratchPath As String = ""

    <AssemblyInitialize>
    Public Shared Sub InitializeScratchStore(context As TestContext)
        _scratchPath = Path.Combine(Path.GetTempPath(), "bta_osg_tests_" & Guid.NewGuid().ToString("N") & ".xml")
        EmbeddedDB.DbPath = _scratchPath
        ' The scratch store is the offline harness every test writes into. Nothing is seeded
        ' any more, so each test enrols the identities it authenticates, and the live-SQL suite
        ' flips SQL mode back on for its own classes.
        AppSettings.Instance.DatabaseSettings.UseSqlServer = False
        EmbeddedDB.Initialize()
    End Sub

    <AssemblyCleanup>
    Public Shared Sub DeleteScratchStore()
        Try
            If _scratchPath.Length > 0 AndAlso File.Exists(_scratchPath) Then File.Delete(_scratchPath)
        Catch
        End Try
    End Sub
End Class

<TestClass>
Public Class OfflineCacheRoundTripTests
    ' Audit-009 CRITICAL: Initialize never read the XML back, so offline work died on
    ' restart. The reload must return the persisted offline rows and keep the identity
    ' lane moving past them.
    <TestMethod>
    Public Sub Initialize_ReloadsPersistedOfflineDocument()
        Dim code = EmbeddedDB.GenerateDocCode("Finance")
        Dim docId = EmbeddedDB.AddDocument(code, "Finance", "Roundtrip persistence probe", "Origin Office", "Finance Section", "CAB-A", "S-1", "BOX-01", "", "RECEIVED", "Hassim A. Ibrahim", isOffline:=True)

        EmbeddedDB.Save()
        EmbeddedDB.Initialize()

        Dim reloaded = EmbeddedDB.DataSet.Tables("Documents").Select("DocumentID = " & docId.ToString())
        Assert.AreEqual(1, reloaded.Length, "offline document must survive a restart")
        Assert.AreEqual("RECEIVED", reloaded(0)("CurrentStatus").ToString())
        Assert.AreEqual(True, CBool(reloaded(0)("PendingSync")), "reloaded offline row must still be flagged for replay")
    End Sub

    <TestMethod>
    Public Sub Initialize_ContinuesIdentityPastReloadedMax()
        Dim firstId = EmbeddedDB.AddDocument(EmbeddedDB.GenerateDocCode("Finance"), "Finance", "Lane probe A", "O", "Finance Section", "CAB-A", "S-1", "BOX-01", "", "RECEIVED", "H", isOffline:=False)
        EmbeddedDB.Save()
        EmbeddedDB.Initialize()
        Dim nextId = EmbeddedDB.AddDocument(EmbeddedDB.GenerateDocCode("Finance"), "Finance", "Lane probe B", "O", "Finance Section", "CAB-A", "S-1", "BOX-01", "", "RECEIVED", "H", isOffline:=False)

        Assert.IsTrue(nextId > firstId, "identity lane must advance past reloaded rows, not reuse ids")
    End Sub

    ' Audit-009 HIGH: offline user enrolment was neither replayed nor protected from the
    ' snapshot doom-sweep. The PendingSync flag is what now keeps the row alive.
    <TestMethod>
    Public Sub OfflineUserEnrolment_SurvivesSnapshotPull()
        EmbeddedDB.AddUser("DEADBEEF01", "Replay Probe User", "Records Section", "Records Section", pendingSync:=True)

        Dim usersTable = EmbeddedDB.DataSet.Tables("Users")
        Dim snapshot = usersTable.Clone()
        Dim mirrored = snapshot.NewRow()
        mirrored("UserID") = 99
        mirrored("RFID_UID") = "MIRROR99"
        mirrored("FullName") = "Mirrored Staff"
        mirrored("Role") = "SG"
        mirrored("Office") = "Office of the Secretary-General"
        snapshot.Rows.Add(mirrored)

        EmbeddedDB.ApplySnapshot("Users", snapshot)

        ' Matched by name: the UID carrier columns are hashed in the cache, and this
        ' test's subject is PendingSync survival, not credential storage.
        Dim survivors = usersTable.Select("FullName = 'Replay Probe User'")
        Assert.AreEqual(1, survivors.Length, "pending offline enrolment must not be wiped by the pull")
        Assert.AreEqual(True, CBool(survivors(0)("PendingSync")))
    End Sub

    ' Audit-009 HIGH: offline free-text statuses replayed to the RECEIVED fallback and
    ' reset approved documents. Plain directives must leave the status column alone.
    <TestMethod>
    Public Sub OfflinePlainDirective_DoesNotCorruptStatus()
        Dim docId = EmbeddedDB.AddDocument(EmbeddedDB.GenerateDocCode("Finance"), "Finance", "Directive status probe", "O", "Finance Section", "CAB-A", "S-1", "BOX-01", "", "FOR_REVIEW", "H", isOffline:=False)

        EmbeddedDB.AddDirective(docId, "For Immediate Action", "Hassim A. Ibrahim", "please expedite", "Prof. Ali B. Pangalian")

        Dim rows = EmbeddedDB.DataSet.Tables("Documents").Select("DocumentID = " & docId.ToString())
        Assert.AreEqual("FOR_REVIEW", rows(0)("CurrentStatus").ToString(), "status column must keep the status code")
        Assert.IsTrue(rows(0)("LastActionTaken").ToString().StartsWith("SG Directive: "), "directive text belongs in last action")
    End Sub

    <TestMethod>
    Public Sub OfflineArchiveDirective_MapsToArchivedStatus()
        Dim docId = EmbeddedDB.AddDocument(EmbeddedDB.GenerateDocCode("Finance"), "Finance", "Archive directive probe", "O", "Finance Section", "CAB-A", "S-1", "BOX-01", "", "FOR_REVIEW", "H", isOffline:=False)

        EmbeddedDB.AddDirective(docId, "Approved & Archived", "Records Section", "approved by SG", "Prof. Ali B. Pangalian", directiveCode:=DesktopDataCoordinator.DirectiveCodeFor("Approved & Archived"))

        Dim rows = EmbeddedDB.DataSet.Tables("Documents").Select("DocumentID = " & docId.ToString())
        Assert.AreEqual("ARCHIVED", rows(0)("CurrentStatus").ToString(), "archive directive must match the SQL ResultStatusID outcome")
    End Sub

    ' Replay fidelity: the audit trail must keep its original action type and the document
    ' row must carry who modified it, not a uniform OFFLINE_SYNC blob.
    <TestMethod>
    Public Sub AuditRows_CarryActionTypeForReplay()
        EmbeddedDB.LogAudit("Probe User", "audit fidelity probe", userId:=7, actionType:="DOCUMENT_CREATED")

        Dim rows = EmbeddedDB.DataSet.Tables("AuditTrail").Select("ActionDescription = 'audit fidelity probe'")
        Assert.AreEqual(1, rows.Length)
        Assert.AreEqual("DOCUMENT_CREATED", rows(0)("ActionType").ToString())
        Assert.AreEqual(7, CInt(rows(0)("UserID")))
    End Sub

    <TestMethod>
    Public Sub DocumentMirror_RecordsModifiedBy()
        Dim docId = EmbeddedDB.AddDocument(EmbeddedDB.GenerateDocCode("Finance"), "Finance", "modified-by probe", "O", "Finance Section", "CAB-A", "S-1", "BOX-01", "", "RECEIVED", "H", isOffline:=True, createdByUserId:=9)
        EmbeddedDB.AddRoutingLog(docId, "O", "Finance Section", "H", "ROUTED", "probe", isOffline:=True, routedByUserId:=12)

        Dim rows = EmbeddedDB.DataSet.Tables("Documents").Select("DocumentID = " & docId.ToString())
        Assert.AreEqual(12, CInt(rows(0)("ModifiedByUserID")), "state replay must attribute the last modifier, not the creator")
    End Sub

    <TestMethod>
    Public Sub DirectiveCodeFor_MapsSeededNames()
        Assert.AreEqual("APPROVE_ARCHIVE", DesktopDataCoordinator.DirectiveCodeFor("Approved & Archived"))
        Assert.AreEqual("REFER_COMMITTEE", DesktopDataCoordinator.DirectiveCodeFor("Referred to Committee on Rules"))
        Assert.AreEqual("IMMEDIATE_ACTION", DesktopDataCoordinator.DirectiveCodeFor("Something Unmapped"))
    End Sub

    <TestMethod>
    Public Sub GenerateDocCode_AdvancesAfterInsert()
        Dim a = EmbeddedDB.GenerateDocCode("Finance")
        Dim docId = EmbeddedDB.AddDocument(a, "Finance", "minting probe", "O", "Finance Section", "CAB-A", "S-1", "BOX-01", "", "RECEIVED", "H", isOffline:=False)
        Dim b = EmbeddedDB.GenerateDocCode("Finance")
        Dim numA As Integer
        Dim numB As Integer
        Integer.TryParse(a.Substring(a.LastIndexOf("-"c) + 1), numA)
        Integer.TryParse(b.Substring(b.LastIndexOf("-"c) + 1), numB)
        Assert.AreEqual(numA + 1, numB, "the next mint after a filed document must advance the counter")
        Assert.AreNotEqual(a, b)
    End Sub

    <TestMethod>
    Public Sub LanDiscovery_Probe_ClosedPort_IsUnreachable()
        ' A silent host must classify as unreachable, not as "no database": the wizard only
        ' prefills servers that answered, so misclassifying dead hosts would fill the fields
        ' with ghosts.
        Dim result = SqlInstanceDiscovery.Probe("localhost", 65534, "", "")
        Assert.AreEqual(SqlInstanceDiscovery.ProbeResult.Unreachable, result)
    End Sub

    ' ---------- zero-scripts setup, offline halves ----------

    <TestMethod>
    Public Sub GoSplitter_SplitsExactMatchBatches()
        ' Scripts 001-015 are plain uppercase GO batches; the splitter feeding both the
        ' provisioner and the live tests must keep batch boundaries and drop trailing air.
        Dim sql = "USE [master];" & vbCrLf & "GO" & vbCrLf &
                  "IF 1 = 1" & vbCrLf & "BEGIN" & vbCrLf & "    SELECT 1;" & vbCrLf & "END" & vbCrLf &
                  "GO" & vbCrLf & vbCrLf
        Dim batches As New List(Of String)(DatabaseProvisioner.SplitBatches(sql))
        Assert.AreEqual(2, batches.Count, "two GO-terminated batches, no trailing empty")
        Assert.IsTrue(batches(0).Contains("USE [master];"), "first batch intact")
        Assert.IsTrue(batches(1).Contains("SELECT 1;"), "second batch intact")
    End Sub

    <TestMethod>
    Public Sub HeartbeatSnapshot_MergesByKeyAndPrunes()
        ' The seats grid reads the Heartbeat mirror; keyed by machine name like the SQL
        ' table, refreshed in place, and a seat that stops reporting is swept once another
        ' seat's heartbeat arrives.
        Dim snapshot = EmbeddedDB.DataSet.Tables("Heartbeat").Clone()
        Dim seatOne = snapshot.NewRow()
        seatOne("MachineName") = "SEAT-ONE"
        seatOne("LastSyncUTC") = "2026-09-28 08:00:00"
        seatOne("AppVersion") = "1.0.0"
        seatOne("PendingOutbox") = 3
        seatOne("LastError") = ""
        snapshot.Rows.Add(seatOne)

        EmbeddedDB.ApplySnapshot("Heartbeat", snapshot)
        Dim rows = EmbeddedDB.DataSet.Tables("Heartbeat").Select("MachineName = 'SEAT-ONE'")
        Assert.AreEqual(1, rows.Length, "heartbeat row merged")
        Assert.AreEqual(3, Convert.ToInt32(rows(0)("PendingOutbox")))

        snapshot = EmbeddedDB.DataSet.Tables("Heartbeat").Clone()
        Dim seatOneRefresh = snapshot.NewRow()
        seatOneRefresh("MachineName") = "SEAT-ONE"
        seatOneRefresh("LastSyncUTC") = "2026-09-28 08:01:00"
        seatOneRefresh("AppVersion") = "1.0.1"
        seatOneRefresh("PendingOutbox") = 0
        snapshot.Rows.Add(seatOneRefresh)
        EmbeddedDB.ApplySnapshot("Heartbeat", snapshot)

        rows = EmbeddedDB.DataSet.Tables("Heartbeat").Select("MachineName = 'SEAT-ONE'")
        Assert.AreEqual(1, rows.Length, "keyed by machine, not duplicated")
        Assert.AreEqual(0, Convert.ToInt32(rows(0)("PendingOutbox")), "refreshed in place")
        Assert.AreEqual("1.0.1", rows(0)("AppVersion").ToString(), "app version follows the seat")

        snapshot = EmbeddedDB.DataSet.Tables("Heartbeat").Clone()
        Dim seatTwo = snapshot.NewRow()
        seatTwo("MachineName") = "SEAT-TWO"
        seatTwo("LastSyncUTC") = "2026-09-28 08:02:00"
        seatTwo("AppVersion") = "1.0.1"
        seatTwo("PendingOutbox") = 0
        snapshot.Rows.Add(seatTwo)
        EmbeddedDB.ApplySnapshot("Heartbeat", snapshot)

        Assert.AreEqual(0, EmbeddedDB.DataSet.Tables("Heartbeat").Select("MachineName = 'SEAT-ONE'").Length,
                        "a seat absent from the pull is pruned like any other mirror row")
        Assert.AreEqual(1, EmbeddedDB.DataSet.Tables("Heartbeat").Select("MachineName = 'SEAT-TWO'").Length)
    End Sub

    <TestMethod>
    Public Sub SyncInterval_FloorHolds()
        Dim s As New DatabaseSettings()
        Assert.AreEqual(10, s.EffectiveSyncIntervalSeconds(), "default interval is 10s")
        s.SyncIntervalSeconds = 0
        Assert.AreEqual(5, s.EffectiveSyncIntervalSeconds(), "a hand-edited zero floors at 5s")
        s.SyncIntervalSeconds = 2
        Assert.AreEqual(5, s.EffectiveSyncIntervalSeconds())
        s.SyncIntervalSeconds = 30
        Assert.AreEqual(30, s.EffectiveSyncIntervalSeconds())
    End Sub
End Class
