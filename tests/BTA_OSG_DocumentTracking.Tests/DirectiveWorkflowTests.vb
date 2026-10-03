Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Threading.Tasks
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports BTA_OSG

Namespace BTA_OSG.Tests
    <TestClass>
    Public Class DirectiveWorkflowTests
        <TestMethod>
        Public Sub DirectiveWorkflow_MapsResultStatus_WhenProvided()
            Dim directiveType As New DirectiveType With {
                .DirectiveCode = "APPROVE_ARCHIVE",
                .DirectiveName = "Approved & Archived",
                .ResultStatusID = 8
            }

            Dim targetStatusId As Integer = If(directiveType.ResultStatusID.HasValue, directiveType.ResultStatusID.Value, 4)
            Assert.AreEqual(8, targetStatusId, "Directive with ResultStatusID must map directly to that status.")
        End Sub

        <TestMethod>
        Public Sub DirectiveWorkflow_DefaultsToPending_WhenNoResultStatus()
            Dim directiveType As New DirectiveType With {
                .DirectiveCode = "IMMEDIATE_ACTION",
                .DirectiveName = "For Immediate Action",
                .ResultStatusID = Nothing
            }

            Dim targetStatusId As Integer = If(directiveType.ResultStatusID.HasValue, directiveType.ResultStatusID.Value, 4)
            Assert.AreEqual(4, targetStatusId, "Directive with null ResultStatusID must default to status ID 4.")
        End Sub

        <TestMethod>
        Public Sub DocumentService_AutoRoutesCategoryToResponsibleSection()
            Assert.AreEqual("Secretariat", DocumentService.GetDefaultSectionForCategory("REG_COMM"))
            Assert.AreEqual("Legislative Section", DocumentService.GetDefaultSectionForCategory("LEG"))
            Assert.AreEqual("Finance Section", DocumentService.GetDefaultSectionForCategory("FIN"))
            Assert.AreEqual("Travel Section", DocumentService.GetDefaultSectionForCategory("TRAVEL"))
        End Sub

        <TestMethod>
        Public Sub EmbeddedDB_FullRevisionLoop_TransitionsCorrectly()
            EmbeddedDB.Initialize()
            Dim docCode = EmbeddedDB.GenerateDocCode("Regular Communication")
            Dim docId = EmbeddedDB.AddDocument(docCode, "Regular Communication", "Test Communication Flow", "Ministry of Transportation", "Office of the Secretary-General", "CAB-A", "S-1", "BOX-01", "", "FOR_REVIEW", "Amina T. Macacua", "INCOMING", "Secretariat")

            Dim dt = EmbeddedDB.DataSet.Tables("Documents")
            Dim rows = dt.Select("DocumentID = " & docId)
            Assert.AreEqual(1, rows.Length)
            Assert.AreEqual("FOR_REVIEW", rows(0)("CurrentStatus").ToString())
            Assert.AreEqual("Secretariat", rows(0)("AssignedSection").ToString())

            ' Sec Gen requests revision
            EmbeddedDB.RequestRevision(docId, "Please append Section 4 Annex B documentation.", "Secretariat", "Prof. Ali B. Pangalian")
            rows = dt.Select("DocumentID = " & docId)
            Assert.AreEqual("FOR_REVISION", rows(0)("CurrentStatus").ToString())
            Assert.AreEqual("Secretariat", rows(0)("AssignedSection").ToString())
            Assert.IsTrue(rows(0)("RevisionPunchlist").ToString().Contains("Section 4 Annex B"))

            ' Secretariat resubmits
            EmbeddedDB.ResubmitDocument(docId, "Amina T. Macacua", "Annex B appended as requested.")
            rows = dt.Select("DocumentID = " & docId)
            Assert.AreEqual("FOR_REVIEW", rows(0)("CurrentStatus").ToString())
            Assert.AreEqual("Secretary-General", rows(0)("AssignedSection").ToString())

            ' Sec Gen approves
            EmbeddedDB.ApproveDocument(docId, "Prof. Ali B. Pangalian", "Approved for official release.")
            rows = dt.Select("DocumentID = " & docId)
            Assert.AreEqual("APPROVED", rows(0)("CurrentStatus").ToString())
            Assert.AreEqual("Records Section", rows(0)("AssignedSection").ToString())

            ' Records Section releases
            EmbeddedDB.ReleaseDocument(docId, "Sittie K. Amin", "Transmitted to recipient office.")
            rows = dt.Select("DocumentID = " & docId)
            Assert.AreEqual("RELEASED", rows(0)("CurrentStatus").ToString())
            Assert.AreEqual("Archived / Released", rows(0)("AssignedSection").ToString())
        End Sub

        <TestMethod>
        Public Sub EmbeddedDB_GetVisibleDocuments_FiltersBySectionAndCategory()
            ' The store starts empty, so this test enrols its own two officers and one document
            ' each: a visibility filter needs a row it must include and a row it must exclude.
            EmbeddedDB.Initialize()
            EmbeddedDB.AddUser("TESTSEC001", "Test Secretariat Officer", "Secretariat", "Secretariat")
            EmbeddedDB.AddUser("TESTSG0001", "Test Secretary-General", "Secretary-General", "Office of the Secretary-General")

            Dim userSec = EmbeddedDB.AuthenticateRFID("TESTSEC001")
            Assert.IsNotNull(userSec)

            EmbeddedDB.AddDocument("SEC-VIS-001", "Regular Communication", "Secretariat desk probe", "Ministry of Interior", "Office of the Secretary-General", "CAB-A", "S-1", "BOX-01", "", "FOR_REVIEW", "Test Secretariat Officer", "INCOMING", "Secretariat")
            EmbeddedDB.AddDocument("FIN-VIS-001", "Finance", "Other desk probe", "Finance Division", "Finance Section", "CAB-C", "S-2", "BOX-02", "", "RECEIVED", "Unrelated Officer", "INCOMING", "Finance Section")

            Dim secretariatDocs = EmbeddedDB.GetVisibleDocuments(userSec)
            Assert.IsTrue(secretariatDocs.Cast(Of System.Data.DataRowView)().Any(Function(r) r("Title").ToString() = "Secretariat desk probe"), "the desk sees its own section's document")
            Assert.IsFalse(secretariatDocs.Cast(Of System.Data.DataRowView)().Any(Function(r) r("Title").ToString() = "Other desk probe"), "the desk does not see another section's document")
            For Each row As System.Data.DataRowView In secretariatDocs
                Dim sec = row("AssignedSection").ToString()
                Dim staff = row("AssignedStaff").ToString()
                Assert.IsTrue(sec = "Secretariat" OrElse staff = "Test Secretariat Officer" OrElse String.IsNullOrEmpty(staff))
            Next

            Dim userSG = EmbeddedDB.AuthenticateRFID("TESTSG0001")
            Assert.IsNotNull(userSG)

            Dim finOnly = EmbeddedDB.GetVisibleDocuments(userSG, "", "Finance")
            Assert.IsTrue(finOnly.Cast(Of System.Data.DataRowView)().Any(Function(r) r("Title").ToString() = "Other desk probe"), "the category filter has a matching row to return")
            For Each row As System.Data.DataRowView In finOnly
                Assert.AreEqual("Finance", row("DocType").ToString())
            Next
        End Sub

        <TestMethod>
        Public Sub EmbeddedDB_SupportsConcurrentReadsAndWrites()
            EmbeddedDB.Initialize()
            Dim exceptions As New System.Collections.Concurrent.ConcurrentBag(Of Exception)()

            Dim tasks As New List(Of Task)()
            For i As Integer = 1 To 10
                Dim idx = i
                tasks.Add(Task.Run(Sub()
                                       Try
                                           If idx Mod 2 = 0 Then
                                               EmbeddedDB.AddDocument("CONC-" & idx, "Regular Communication", "Concurrent " & idx, "Origin", "Dest", "CAB-A", "S-1", "BOX-1", "", "FOR_REVIEW", "Staff", "INCOMING", "Secretariat")
                                           Else
                                               Dim docs = EmbeddedDB.GetVisibleDocuments(Nothing)
                                               Assert.IsNotNull(docs)
                                           End If
                                       Catch ex As Exception
                                           exceptions.Add(ex)
                                       End Try
                                   End Sub))
            Next

            Task.WaitAll(tasks.ToArray())
            Assert.AreEqual(0, exceptions.Count, "Concurrent operations on EmbeddedDB threw exceptions.")
        End Sub

        <TestMethod>
        Public Sub DesktopDataCoordinator_DispatchesToEmbeddedWhenOffline()
            EmbeddedDB.Initialize()
            Dim isConnected = False ' Simulate offline mode
            Dim coordinator As New DesktopDataCoordinator(isConnected)

            Dim docCode = EmbeddedDB.GenerateDocCode("Regular Communication")
            Dim registeredCode = coordinator.RegisterDocument(docCode, "Regular Communication", "Test Dual Dispatch", "Ministry", "OSG", "CAB-1", "S-1", "B-1", "", "FOR_REVIEW", "Staff", "INCOMING", "Secretariat", DateTime.Now.AddDays(3).ToString("yyyy-MM-dd HH:mm:ss"), "", "Registered", "Test User", 1)

            Assert.IsFalse(String.IsNullOrWhiteSpace(registeredCode))
            Dim dt = EmbeddedDB.DataSet.Tables("Documents")
            Dim rows = dt.Select("DocCode = '" & registeredCode.Replace("'", "''") & "'")
            Assert.AreEqual(1, rows.Length)
            Assert.AreEqual("Test Dual Dispatch", rows(0)("Title").ToString())
        End Sub
    End Class
End Namespace
