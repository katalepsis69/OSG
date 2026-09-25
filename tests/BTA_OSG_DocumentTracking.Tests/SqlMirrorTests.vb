Option Explicit On
Option Strict On

Imports System
Imports System.Data
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports BTA_OSG

Namespace BTA_OSG.Tests
    ''' <summary>
    ''' The cache mirror is what makes two workstations agree, and its load-bearing
    ''' behaviours (identity values survive the snapshot; the auto-increment lane moves past
    ''' them) are DataSet details, not SQL details, so they are testable without a server.
    ''' </summary>
    <TestClass>
    Public Class SqlMirrorTests
        Private Shared Function BuildSnapshot(ParamArray ids As Integer()) As DataTable
            Dim source = EmbeddedDB.DataSet.Tables("Documents").Clone()
            For Each id In ids
                Dim row = source.NewRow()
                row("DocumentID") = id
                row("DocCode") = "TST-2026-" & id.ToString()
                row("DocType") = "Regular Communication"
                row("Title") = "Mirror probe " & id.ToString()
                row("CurrentStatus") = "RECEIVED"
                row("FlowDirection") = "INCOMING"
                row("AssignedSection") = "Records Section"
                row("DateReceived") = "2026-09-25 08:00:00"
                source.Rows.Add(row)
            Next
            Return source
        End Function

        <TestMethod>
        Public Sub ApplySnapshot_ReplacesTheTableWithSqlRows()
            EmbeddedDB.Initialize()
            Dim docs = EmbeddedDB.DataSet.Tables("Documents")
            docs.Rows.Clear()

            ' A stale row the snapshot no longer lists must leave the table, or a document
            ' deleted in SQL would linger in every grid forever.
            EmbeddedDB.ApplySnapshot("Documents", BuildSnapshot(9101, 9102, 9103))
            EmbeddedDB.ApplySnapshot("Documents", BuildSnapshot(9001, 9002))

            Assert.AreEqual(2, docs.Rows.Count, "the snapshot replaces whatever the cache held")
            Assert.IsNull(docs.Rows.Find(9101), "rows absent from the snapshot are dropped")
            Assert.IsNotNull(docs.Rows.Find(9001), "SQL identity values are the cache keys")
            Assert.AreEqual("TST-2026-9001", docs.Rows.Find(9001)("DocCode").ToString())
        End Sub

        <TestMethod>
        Public Sub ApplySnapshot_KeepsTheNextLocalIdClearOfMirroredIds()
            EmbeddedDB.Initialize()
            ' Ids a fresh SQL Server actually mints: 1, 2, 3. High ids would let a lane that
            ' never advanced look healthy, because the seed it fell back to (1) would still be
            ' free. These are the ids that collide when the lane does not follow the snapshot.
            EmbeddedDB.ApplySnapshot("Documents", BuildSnapshot(1, 2, 3))

            Dim localId = EmbeddedDB.AddDocument("LOC-2026-001", "Finance", "Offline registration", "OSG", "Finance Section", "CAB-A", "S-1", "BOX-01", "", "RECEIVED", "Hassim A. Ibrahim", "INCOMING", "Finance Section")
            Assert.AreNotEqual(1, localId, "a local row must not reuse a mirrored id or the primary key constraint throws")
            Assert.AreNotEqual(2, localId)
            Assert.AreNotEqual(3, localId)
        End Sub

        <TestMethod>
        Public Sub ApplySnapshot_UpdatesHeldRowsInPlace()
            EmbeddedDB.Initialize()
            EmbeddedDB.ApplySnapshot("Documents", BuildSnapshot(9011))
            Dim held = EmbeddedDB.DataSet.Tables("Documents").Rows.Find(9011)
            Assert.IsNotNull(held)

            ' The same document, edited at another workstation: a held row must follow it.
            Dim edited = BuildSnapshot(9011, 9021)
            edited.Rows.Find(9011)("Title") = "Edited at another desk"
            EmbeddedDB.ApplySnapshot("Documents", edited)

            Assert.AreSame(held, EmbeddedDB.DataSet.Tables("Documents").Rows.Find(9011), "rows are updated in place, not detached")
            Assert.AreEqual("Edited at another desk", held("Title").ToString(), "the held row sees the mirrored values")
            Assert.IsNotNull(held.Table, "FormDocumentDetail.DocRow is dereferenced for as long as the dialog is open")
            Assert.IsTrue(held.Table.Columns.Contains("DocCode"))
        End Sub

        <TestMethod>
        Public Sub ApplySnapshot_PreservesPendingOfflineRows()
            EmbeddedDB.Initialize()
            Dim docs = EmbeddedDB.DataSet.Tables("Documents")
            docs.Rows.Clear()

            ' Register an offline document (tagged PendingSync = True)
            Dim localId = EmbeddedDB.AddDocument("LOC-PENDING-001", "Finance", "Offline Budget Proposal", "OSG", "Finance Section", "CAB-A", "S-1", "BOX-01", "", "RECEIVED", "Hassim A. Ibrahim", "INCOMING", "Finance Section")
            Dim localRow = docs.Rows.Find(localId)
            Assert.IsNotNull(localRow)
            Assert.IsTrue(CBool(localRow("PendingSync")), "Offline document must be tagged PendingSync = True")

            ' Apply snapshot from SQL Server that does NOT have this local document yet
            Dim serverSnapshot = BuildSnapshot(9001, 9002)
            EmbeddedDB.ApplySnapshot("Documents", serverSnapshot)

            ' Assert the offline document was NOT deleted
            Dim survivingLocalRow = docs.Rows.Find(localId)
            Assert.IsNotNull(survivingLocalRow, "Offline rows marked PendingSync must never be wiped by SQL snapshot pulls")
            Assert.AreEqual("LOC-PENDING-001", survivingLocalRow("DocCode").ToString())

            ' Assert the server rows were also loaded
            Assert.IsNotNull(docs.Rows.Find(9001))
            Assert.IsNotNull(docs.Rows.Find(9002))
        End Sub

        <TestMethod>
        Public Sub BuildSqlSnapshot_DisconnectedHost_ReturnsNothingWithoutThrowing()
            Dim previous = Program.IsDatabaseConnected
            Try
                Program.IsDatabaseConnected = False
                Assert.IsNull(New DesktopDataCoordinator(False).BuildSqlSnapshot(), "the poller must no-op, not throw, on a disconnected workstation")
                Assert.IsNull(New DesktopDataCoordinator(True).BuildSqlSnapshot())
            Finally
                Program.IsDatabaseConnected = previous
            End Try
        End Sub

        <TestMethod>
        Public Sub SyncOfflineOutbox_Disconnected_ReturnsZeroWithoutThrowing()
            Dim previous = Program.IsDatabaseConnected
            Try
                Program.IsDatabaseConnected = False
                Dim coordinator As New DesktopDataCoordinator(False)
                Dim syncedCount = coordinator.SyncOfflineOutbox()
                Assert.AreEqual(0, syncedCount, "Disconnected outbox sync must return 0 without throwing")
            Finally
                Program.IsDatabaseConnected = previous
            End Try
        End Sub

        <TestMethod>
        Public Sub RouteDocument_Offline_UpdatesDestinationAndLastAction()
            EmbeddedDB.Initialize()
            Dim localId = EmbeddedDB.AddDocument("LOC-ROUTE-001", "Finance", "Routing Check", "Records Section", "Finance Section", "CAB-A", "S-1", "BOX-01", "", "RECEIVED", "Hassim A. Ibrahim", "INCOMING", "Finance Section")

            Dim coordinator As New DesktopDataCoordinator(False)
            coordinator.RouteDocument(localId, "Finance Section", "Office of the Secretary-General", "FOR_REVIEW", "Please review budget items", "Hassim A. Ibrahim", 4)

            Dim docRow = EmbeddedDB.DataSet.Tables("Documents").Rows.Find(localId)
            Assert.IsNotNull(docRow)
            Assert.AreEqual("Office of the Secretary-General", docRow("DestinationOffice").ToString())
            Assert.IsTrue(docRow("LastActionTaken").ToString().Contains("Office of the Secretary-General"), "LastActionTaken must reflect destination office")
            Assert.IsTrue(docRow("LastActionTaken").ToString().Contains("FOR_REVIEW"), "LastActionTaken must reflect the routing action")
        End Sub
    End Class
End Namespace
