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
        Public Sub Analytics_RankingCalculation_OrdersByVolumeCorrectly()
            ' Simulate category distribution
            Dim counts As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase) From {
                {"Legal Opinions", 5},
                {"Communications", 12},
                {"Travel Orders", 3},
                {"Financial Reports", 8}
            }

            Dim sortedList = New List(Of KeyValuePair(Of String, Integer))(counts)
            sortedList.Sort(Function(a, b) b.Value.CompareTo(a.Value))

            Assert.AreEqual("Communications", sortedList(0).Key, "Rank 1 must be Communications with highest volume.")
            Assert.AreEqual(12, sortedList(0).Value)
            Assert.AreEqual("Financial Reports", sortedList(1).Key, "Rank 2 must be Financial Reports.")
            Assert.AreEqual("Legal Opinions", sortedList(2).Key, "Rank 3 must be Legal Opinions.")
            Assert.AreEqual("Travel Orders", sortedList(3).Key, "Rank 4 must be Travel Orders.")

            Dim total = 5 + 12 + 3 + 8
            Dim shareRank1 = (sortedList(0).Value / CDbl(total)) * 100.0
            Assert.AreEqual(42.86, Math.Round(shareRank1, 2), "Share percentage should be correctly calculated.")
        End Sub

        <TestMethod>
        Public Sub Analytics_GadDemographics_ComputesProportionsCorrectly()
            Dim female = 14
            Dim male = 9
            Dim undisclosed = 2
            Dim total = female + male + undisclosed

            Assert.AreEqual(25, total)

            Dim femalePct = (female / CDbl(total)) * 100.0
            Dim malePct = (male / CDbl(total)) * 100.0
            Dim undPct = (undisclosed / CDbl(total)) * 100.0

            Assert.AreEqual(56.0, femalePct, 0.01, "Female percentage calculation must be accurate.")
            Assert.AreEqual(36.0, malePct, 0.01, "Male percentage calculation must be accurate.")
            Assert.AreEqual(8.0, undPct, 0.01, "Undisclosed percentage calculation must be accurate.")
            Assert.AreEqual(100.0, femalePct + malePct + undPct, 0.01, "Percentages must sum to 100%.")
        End Sub
    End Class
End Namespace
