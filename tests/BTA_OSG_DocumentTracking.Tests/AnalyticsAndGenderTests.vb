Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Data
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports BTA_OSG

Namespace BTA_OSG.Tests
    <TestClass>
    Public Class AnalyticsAndGenderTests
        <TestInitialize>
        Public Sub Setup()
            EmbeddedDB.Initialize()
        End Sub

        <TestMethod>
        Public Sub PortalSubmission_GenderProperty_StoresAndRetrievesValues()
            Dim sub1 As New PortalSubmission With {
                .ControlNumber = "COMM-2026-0001-ABCD",
                .Category = "COMM",
                .DocumentTitle = "Test Document",
                .RequesterName = "Fatima Alawi",
                .RequesterEmail = "fatima@example.gov.ph",
                .RequesterPhone = "09171234567",
                .RequesterGender = "Female",
                .CreatedAt = DateTime.UtcNow
            }

            Assert.AreEqual("Female", sub1.RequesterGender, "RequesterGender must carry the set value.")

            Dim sub2 As New PortalSubmission()
            Assert.AreEqual("", sub2.RequesterGender, "RequesterGender must default to empty string.")
        End Sub

        <TestMethod>
        Public Sub EmbeddedDB_DocumentsTable_HasRequesterGenderColumn()
            Dim dt = EmbeddedDB.DataSet.Tables("Documents")
            Assert.IsNotNull(dt, "Documents table must exist in EmbeddedDB.")
            Assert.IsTrue(dt.Columns.Contains("RequesterGender"), "Documents table must contain RequesterGender column for GAD compliance.")
        End Sub

        <TestMethod>
        Public Sub EmbeddedDB_AddDocument_StoresRequesterGenderCorrectly()
            Dim docId = EmbeddedDB.AddDocument(
                code:="COMM-2026-9999",
                docType:="COMM",
                title:="GAD Demographic Test Doc",
                origin:="Amina Usman (amina@bta.gov.ph)",
                dest:="Office of the Secretary-General",
                cab:="CAB-A", shelf:="S-1", box:="BOX-01",
                url:="",
                status:="RECEIVED",
                assigned:="Omire Khalid B. Ebrahim",
                flowDirection:="INCOMING",
                assignedSection:="Secretariat",
                targetDeadline:="",
                punchlist:="",
                lastAction:="Registered for GAD test",
                externalControlNumber:="COMM-2026-9999-WXYZ",
                isOffline:=True,
                createdByUserId:=1,
                requesterGender:="Female")

            Dim rows = EmbeddedDB.DataSet.Tables("Documents").Select("DocumentID = " & docId)
            Assert.IsTrue(rows.Length > 0, "Document must be found in EmbeddedDB.")
            Assert.AreEqual("Female", rows(0)("RequesterGender").ToString(), "RequesterGender must match stored value.")
            Assert.AreEqual("COMM-2026-9999-WXYZ", rows(0)("ExternalControlNumber").ToString(), "ExternalControlNumber must match.")
        End Sub

        <TestMethod>
        Public Sub CollectAnalytics_CountsStatusesCategoriesAndGender_FromRealRows()
            ' Drives the analytics engine's own computation against a seeded table, so a
            ' regression in the aggregation, ranking, or demographics math fails here.
            Dim dt As New DataTable("Documents")
            dt.Columns.Add("DateReceived", GetType(String))
            dt.Columns.Add("AssignedSection", GetType(String))
            dt.Columns.Add("CurrentStatus", GetType(String))
            dt.Columns.Add("DocType", GetType(String))
            dt.Columns.Add("RequesterGender", GetType(String))

            Dim seed = New List(Of String()) From {
                New String() {DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd HH:mm:ss"), "Finance Section", "APPROVED", "Finance", "Female"},
                New String() {DateTime.Now.AddDays(-2).ToString("yyyy-MM-dd HH:mm:ss"), "Finance Section", "RELEASED", "Finance", "Male"},
                New String() {DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd HH:mm:ss"), "Secretariat", "FOR_REVIEW", "Regular Communication", "Female"},
                New String() {DateTime.Now.AddDays(-3).ToString("yyyy-MM-dd HH:mm:ss"), "Secretariat", "FOR_REVISION", "Legislative", "Male"}
            }
            For Each values In seed
                dt.Rows.Add(values(0), values(1), values(2), values(3), values(4))
            Next

            Dim main As New FormMainAnalyticsProbe()
            Dim results = main.ProbeCollectAnalytics(dt, DateTime.MinValue)

            Assert.AreEqual(4, results.MatchedRows.Count, "all seeded rows must be matched")
            Assert.AreEqual(2, results.TotalApproved, "two seeded rows are APPROVED/RELEASED")
            Assert.AreEqual(1, results.TotalPending, "the engine counts revision/reject rows as action-required; FOR_REVIEW is its own state")
            Assert.AreEqual(2, results.CategoryCounts("Finance"), "the Finance category counts its rows")
            Assert.AreEqual(2, results.FemaleCount, "gender demographics count from the RequesterGender column")
            Assert.AreEqual(2, results.MaleCount)
            Assert.AreEqual(2, results.SectionStats("Finance Section").TotalAssigned, "section stats aggregate per desk")
            Assert.IsTrue(results.AvgVelocityDays > 0 AndAlso results.AvgVelocityDays <= 90.0, "average document age stays in the capped window")
        End Sub
    End Class

    ''' <summary>
    ''' CollectAnalytics is an instance method on the main form's analytics partial; this
    ''' constructible probe exposes it without instantiating the full shell.
    ''' </summary>
    Friend Class FormMainAnalyticsProbe
        Inherits FormMain
        Public Sub New()
            MyBase.New()
        End Sub

        Public Function ProbeCollectAnalytics(dtDocs As DataTable, thresholdDate As DateTime) As AnalyticsSummary
            Return CollectAnalytics(dtDocs, thresholdDate)
        End Function
    End Class
End Namespace
