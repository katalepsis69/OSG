Option Explicit On
Option Strict On

Imports System
Imports System.Data
Imports System.IO
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports BTA_OSG

Namespace BTA_OSG.Tests
    ''' <summary>
    ''' Production hardening probes: cache durability, live SQL gating, void eligibility,
    ''' time-basis conversion, and the persisted terminal lockout. All of these are cache or
    ''' pure-function behaviours, so they run without a SQL Server.
    ''' </summary>
    <TestClass>
    Public Class ProductionQolTests

        Private Shared Sub DropTestDocument(docId As Integer)
            SyncLock EmbeddedDB.SyncRoot
                For Each name In {"Documents", "Directives", "RoutingLogs", "Movements"}
                    Dim table = EmbeddedDB.DataSet.Tables(name)
                    If table Is Nothing OrElse Not table.Columns.Contains("DocumentID") Then Continue For
                    For Each row As DataRow In table.Select("DocumentID = " & docId.ToString())
                        row.Delete()
                    Next
                Next
            End SyncLock
        End Sub

        <TestMethod>
        Public Sub Cache_QuarantinesCorruptFileAndRecoversPreviousSave()
            Dim previousPath = EmbeddedDB.DbPath
            Dim scratch = Path.Combine(Path.GetTempPath(), "bta_qol_" & Guid.NewGuid().ToString("N") & ".xml")
            Try
                EmbeddedDB.DbPath = scratch
                EmbeddedDB.Initialize()
                EmbeddedDB.AddDocument("QOL-001", "Finance", "Recovered row", "O", "Finance Section", "CAB-A", "S-1", "BOX-01", "", "RECEIVED", "H", isOffline:=False)
                EmbeddedDB.Save()
                EmbeddedDB.AddDocument("QOL-002", "Finance", "Lost row", "O", "Finance Section", "CAB-A", "S-1", "BOX-02", "", "RECEIVED", "H", isOffline:=False)
                EmbeddedDB.Save()

                Assert.IsTrue(File.Exists(scratch & ".bak"), "the second save must keep the previous generation")
                File.WriteAllText(scratch, "<not-a-dataset")
                EmbeddedDB.Initialize()

                Assert.IsTrue(EmbeddedDB.LastLoadError.Contains(".corrupt-"), "the unreadable file must be kept aside")
                Dim docs = EmbeddedDB.DataSet.Tables("Documents")
                Assert.AreEqual(1, docs.Select("Title = 'Recovered row'").Length, "the .bak generation must be restored")
                Assert.AreEqual(0, docs.Select("Title = 'Lost row'").Length, "the unreadable generation cannot be recovered")
                Dim quarantined = Directory.GetFiles(Path.GetDirectoryName(scratch), Path.GetFileName(scratch) & ".corrupt-*")
                Assert.AreEqual(1, quarantined.Length, "the corrupt file is retained for recovery")
            Finally
                EmbeddedDB.DbPath = previousPath
                EmbeddedDB.Initialize()
                For Each leftover In Directory.GetFiles(Path.GetDirectoryName(scratch), Path.GetFileName(scratch) & "*")
                    Try
                        File.Delete(leftover)
                    Catch
                    End Try
                Next
            End Try
        End Sub

        <TestMethod>
        Public Sub Coordinator_OfflineConstructionIgnoresStaleGlobalConnection()
            Dim previous = Program.IsDatabaseConnected
            Try
                Program.IsDatabaseConnected = True
                Dim offline As New DesktopDataCoordinator(False)
                Assert.IsFalse(offline.IsDatabaseConnected, "a seat constructed offline must not write SQL because a global flag is stale")
                Dim online As New DesktopDataCoordinator(True)
                Assert.IsTrue(online.IsDatabaseConnected)
            Finally
                Program.IsDatabaseConnected = previous
            End Try
        End Sub

        <TestMethod>
        Public Sub VoidDocument_RemovesLocalOnlyRowAndAudits()
            EmbeddedDB.Initialize()
            Dim code = "VOID-" & Guid.NewGuid().ToString("N").Substring(0, 8).ToUpperInvariant()
            Dim docId = EmbeddedDB.AddDocument(code, "Finance", "Void probe", "O", "Finance Section", "CAB-V", "S-1", "BOX-V", "", "RECEIVED", "H", isOffline:=True)
            Dim coordinator As New DesktopDataCoordinator(False)

            Assert.IsTrue(coordinator.VoidDocument(docId, "Wrong document scanned at intake", "Test Admin", 1, "System Administrator"))
            Assert.AreEqual(0, EmbeddedDB.DataSet.Tables("Documents").Select("DocumentID = " & docId.ToString()).Length, "the local row is removed")
            Dim auditRows = EmbeddedDB.DataSet.Tables("AuditTrail").Select("ActionType = 'DOCUMENT_VOIDED'")
            Assert.AreEqual(1, auditRows.Length, "the void is audited")
            Assert.IsTrue(auditRows(0)("ActionDescription").ToString().Contains(code))
        End Sub

        <TestMethod>
        Public Sub VoidDocument_RefusesOfflineSyncedAndTerminalRows()
            EmbeddedDB.Initialize()
            Dim syncedId As Integer = 0
            Dim terminalId As Integer = 0
            Try
                syncedId = EmbeddedDB.AddDocument("SYNC-" & Guid.NewGuid().ToString("N").Substring(0, 8).ToUpperInvariant(), "Finance", "Synced void probe", "O", "Finance Section", "CAB-S", "S-1", "BOX-S", "", "RECEIVED", "H", isOffline:=False)
                terminalId = EmbeddedDB.AddDocument("TERM-" & Guid.NewGuid().ToString("N").Substring(0, 8).ToUpperInvariant(), "Finance", "Terminal void probe", "O", "Finance Section", "CAB-T", "S-1", "BOX-T", "", "RELEASED", "H", isOffline:=True)
                Dim coordinator As New DesktopDataCoordinator(False)

                Assert.IsFalse(coordinator.VoidDocument(syncedId, "Offline void attempt", "Test Admin", 1, "System Administrator"), "a synced document cannot be voided offline")
                Assert.AreEqual(1, EmbeddedDB.DataSet.Tables("Documents").Select("DocumentID = " & syncedId.ToString()).Length)
                Assert.IsFalse(coordinator.VoidDocument(terminalId, "Terminal void attempt", "Test Admin", 1, "System Administrator"), "a terminal document cannot be voided")
                Assert.IsFalse(coordinator.VoidDocument(terminalId, "Unauthorized void attempt", "Desk Clerk", 1, "Records Section"), "a non-owner role cannot void")
                Assert.AreEqual(1, EmbeddedDB.DataSet.Tables("Documents").Select("DocumentID = " & terminalId.ToString()).Length)
            Finally
                If syncedId > 0 Then DropTestDocument(syncedId)
                If terminalId > 0 Then DropTestDocument(terminalId)
            End Try
        End Sub

        <TestMethod>
        Public Sub ApplySnapshot_RemovesAVoidedEarlyDocumentThatWasMirrored()
            EmbeddedDB.Initialize()
            Dim docs = EmbeddedDB.DataSet.Tables("Documents")
            docs.Rows.Clear()

            Dim snapshot = docs.Clone()
            Dim row = snapshot.NewRow()
            row("DocumentID") = 2
            row("DocCode") = "EARLY-2"
            row("Title") = "Early mirrored row"
            row("CurrentStatus") = "RECEIVED"
            row("FlowDirection") = "INCOMING"
            row("RowVersion") = New Byte() {1, 2, 3}
            snapshot.Rows.Add(row)
            EmbeddedDB.ApplySnapshot("Documents", snapshot)
            Assert.IsNotNull(docs.Rows.Find(2))

            ' SQL now omits it after a void. RowVersion proves it is a mirrored row, so the
            ' local ID 1-4 seed guard must not keep it on the seat.
            Dim nextSnapshot = docs.Clone()
            Dim other = nextSnapshot.NewRow()
            other("DocumentID") = 99
            other("DocCode") = "KEEP-99"
            other("Title") = "Keeps the snapshot non-empty"
            nextSnapshot.Rows.Add(other)
            EmbeddedDB.ApplySnapshot("Documents", nextSnapshot)
            Assert.IsNull(docs.Rows.Find(2), "a voided early identity must not be held by the seed guard")
        End Sub

        <TestMethod>
        Public Sub CachedEventUtc_ConvertsWallClockToUtc()
            Dim table As New DataTable("Directives")
            table.Columns.Add("Timestamp", GetType(String))
            Dim wallClock As New DateTime(2026, 3, 4, 15, 30, 0, DateTimeKind.Unspecified)
            Dim row = table.NewRow()
            row("Timestamp") = wallClock.ToString("yyyy-MM-dd HH:mm:ss")
            table.Rows.Add(row)

            Dim expected = DateTime.SpecifyKind(wallClock, DateTimeKind.Local).ToUniversalTime()
            Assert.AreEqual(expected, DesktopDataCoordinator.CachedEventUtc(row, "Timestamp"))
        End Sub

        <TestMethod>
        Public Sub NormalizeMirroredTimes_ConvertsLogsButLeavesHeartbeatUtc()
            Dim logs As New DataTable("Directives")
            logs.Columns.Add("Timestamp", GetType(String))
            logs.Rows.Add("2026-01-02 03:04:05")
            DesktopDataCoordinator.NormalizeMirroredTimes("Directives", logs)

            Dim expected = DateTime.SpecifyKind(New DateTime(2026, 1, 2, 3, 4, 5), DateTimeKind.Utc).ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss")
            Assert.AreEqual(expected, logs.Rows(0)("Timestamp").ToString())

            Dim heartbeat As New DataTable("Heartbeat")
            heartbeat.Columns.Add("LastSyncUTC", GetType(String))
            heartbeat.Rows.Add("2026-01-02 03:04:05")
            DesktopDataCoordinator.NormalizeMirroredTimes("Heartbeat", heartbeat)
            Assert.AreEqual("2026-01-02 03:04:05", heartbeat.Rows(0)("LastSyncUTC").ToString(), "staleness math compares this value with UtcNow")
        End Sub

        <TestMethod>
        Public Sub TerminalLockout_ReadsPersistedExpiry()
            Dim previousPath = EmbeddedDB.LockoutFilePath
            Dim scratch = Path.Combine(Path.GetTempPath(), "bta_lockout_" & Guid.NewGuid().ToString("N") & ".txt")
            Try
                EmbeddedDB.LockoutFilePath = scratch
                EmbeddedDB.ResetTerminalLockout()
                File.WriteAllText(scratch, DateTime.UtcNow.AddMinutes(4).ToString("o"))
                Assert.IsTrue(EmbeddedDB.IsTerminalLockedOut(), "a restart must not clear an unexpired lockout")

                EmbeddedDB.ResetTerminalLockout()
                File.WriteAllText(scratch, DateTime.UtcNow.AddMinutes(-1).ToString("o"))
                Assert.IsFalse(EmbeddedDB.IsTerminalLockedOut(), "an expired lockout releases the terminal")
            Finally
                EmbeddedDB.ResetTerminalLockout()
                EmbeddedDB.LockoutFilePath = previousPath
                Try
                    If File.Exists(scratch) Then File.Delete(scratch)
                Catch
                End Try
            End Try
        End Sub

    End Class
End Namespace
