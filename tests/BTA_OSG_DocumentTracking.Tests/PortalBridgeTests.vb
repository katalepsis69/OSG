Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports BTA_OSG

Namespace BTA_OSG.Tests
    ''' <summary>
    ''' Fake in-memory test double for IPortalBridge to test isolated bridge integration.
    ''' </summary>
    Public Class FakePortalBridge
        Implements IPortalBridge

        Public ReadOnly Queue As New List(Of PortalSubmission)()
        Public ReadOnly AcknowledgedNumbers As New List(Of String)()
        Public ReadOnly StatusPushes As New Dictionary(Of String, String)()

        Public Function FetchExternalQueueAsync() As Task(Of List(Of PortalSubmission)) Implements IPortalBridge.FetchExternalQueueAsync
            Dim result = New List(Of PortalSubmission)(Queue)
            Return Task.FromResult(result)
        End Function

        Public Function AcknowledgeImportAsync(controlNumber As String) As Task(Of Boolean) Implements IPortalBridge.AcknowledgeImportAsync
            AcknowledgedNumbers.Add(controlNumber)
            Queue.RemoveAll(Function(item) item.ControlNumber = controlNumber)
            Return Task.FromResult(True)
        End Function

        Public Function AcknowledgeBatchImportAsync(controlNumbers As IEnumerable(Of String)) As Task(Of Boolean) Implements IPortalBridge.AcknowledgeBatchImportAsync
            If controlNumbers IsNot Nothing Then
                For Each cn In controlNumbers
                    AcknowledgedNumbers.Add(cn)
                    Queue.RemoveAll(Function(item) item.ControlNumber = cn)
                Next
            End If
            Return Task.FromResult(True)
        End Function

        Public Function PushPublicStatusAsync(controlNumber As String, publicStatus As String) As Task(Of Boolean) Implements IPortalBridge.PushPublicStatusAsync
            StatusPushes(controlNumber) = publicStatus
            Return Task.FromResult(True)
        End Function
    End Class

    <TestClass>
    Public Class PortalBridgeTests
        <TestMethod>
        Public Sub PortalSettings_Defaults_AreSafeAndOffline()
            Dim settings As New PortalSettings()
            Assert.IsFalse(settings.PortalEnabled, "Portal must be disabled by default for offline safety.")
            Assert.IsTrue(settings.TimeoutSeconds >= 3, "Timeout should be at least 3 seconds.")
            Assert.IsFalse(String.IsNullOrWhiteSpace(settings.BaseUrl))
        End Sub

        <TestMethod>
        Public Async Function PortalBridge_WhenDisabled_ReturnsSafeFallbacksWithoutNetwork() As Task
            Dim settings As New PortalSettings With {
                .PortalEnabled = False,
                .BaseUrl = "http://127.0.0.1:65534"
            }
            Dim bridge As New PortalBridge(settings)

            Dim queue = Await bridge.FetchExternalQueueAsync()
            Assert.IsNotNull(queue)
            Assert.AreEqual(0, queue.Count)

            Dim ack = Await bridge.AcknowledgeImportAsync("COMM-2026-0001-K9X2")
            Assert.IsTrue(ack, "Disabled bridge returns true immediately for ack.")

            Dim push = Await bridge.PushPublicStatusAsync("COMM-2026-0001-K9X2", "Approved")
            Assert.IsTrue(push, "Disabled bridge returns true immediately for status push.")
        End Function

        <TestMethod>
        Public Async Function PortalBridge_AcknowledgeBatchImportAsync_ReturnsTrueOnDisabledPortal() As Task
            Dim settings As New PortalSettings With {.PortalEnabled = False}
            Dim bridge As New PortalBridge(settings)
            Dim result = Await bridge.AcknowledgeBatchImportAsync(New List(Of String) From {"CN-001", "CN-002"})
            Assert.IsTrue(result)
        End Function

        <TestMethod>
        Public Async Function PortalBridge_WhenServerUnreachable_SwallowsExceptionAndReturnsFallback() As Task
            Dim settings As New PortalSettings With {
                .PortalEnabled = True,
                .BaseUrl = "http://127.0.0.1:65534",
                .TimeoutSeconds = 1
            }
            Dim bridge As New PortalBridge(settings)

            ' Must never throw unhandled exception even if endpoint is down
            Dim queue = Await bridge.FetchExternalQueueAsync()
            Assert.IsNotNull(queue)
            Assert.AreEqual(0, queue.Count)

            Dim ack = Await bridge.AcknowledgeImportAsync("COMM-2026-0001-K9X2")
            Assert.IsFalse(ack)

            Dim push = Await bridge.PushPublicStatusAsync("COMM-2026-0001-K9X2", "Approved")
            Assert.IsFalse(push)
        End Function

        <TestMethod>
        Public Sub DocumentService_MapToPublicStatus_CorrectlyProjectsInternalStatuses()
            ' Verification against Spec Section 7 Public Mapping Table
            Assert.AreEqual("Received", DocumentService.MapToPublicStatus("RECEIVED"))
            Assert.AreEqual("Under Review", DocumentService.MapToPublicStatus("FOR_REVIEW"))
            Assert.AreEqual("Under Review", DocumentService.MapToPublicStatus("PENDING_REVIEW"))
            Assert.AreEqual("For Processing", DocumentService.MapToPublicStatus("FOR_REVISION"))
            Assert.AreEqual("For Processing", DocumentService.MapToPublicStatus("REVISION_REQUESTED"))
            Assert.AreEqual("Approved", DocumentService.MapToPublicStatus("APPROVED"))
            Assert.AreEqual("Ready for Release", DocumentService.MapToPublicStatus("RELEASED"))
            Assert.AreEqual("Ready for Release", DocumentService.MapToPublicStatus("FILED"))
            Assert.AreEqual("Rejected", DocumentService.MapToPublicStatus("REJECTED"))
            Assert.AreEqual("Rejected", DocumentService.MapToPublicStatus("CANCELLED"))
            Assert.AreEqual("Received", DocumentService.MapToPublicStatus(Nothing))
            Assert.AreEqual("Received", DocumentService.MapToPublicStatus("UNKNOWN_CODE"))
        End Sub

        <TestMethod>
        Public Sub ControlNumber_Format_ValidatesSpecGrammar()
            ' Pattern: PREFIX-YYYY-0000-XXXX
            ' Alphabet: 31 uppercase chars excluding I, L, O, 0, 1
            Dim pattern = "^(COMM|LEG|FIN|TO)-\d{4}-\d{4}-[ABCDEFGHJKMNPQRSTUVWXYZ23456789]{4}$"

            Assert.IsTrue(Regex.IsMatch("COMM-2026-0001-K9X2", pattern))
            Assert.IsTrue(Regex.IsMatch("LEG-2026-0042-3M8P", pattern))
            Assert.IsTrue(Regex.IsMatch("FIN-2026-0100-W7V4", pattern))
            Assert.IsTrue(Regex.IsMatch("TO-2026-9999-2222", pattern))

            ' Invalid cases (wrong prefix, contains forbidden chars 0, O, 1, I, L)
            Assert.IsFalse(Regex.IsMatch("XYZ-2026-0001-K9X2", pattern))
            Assert.IsFalse(Regex.IsMatch("COMM-2026-1-K9X2", pattern))
            Assert.IsFalse(Regex.IsMatch("COMM-2026-0001-K0X2", pattern)) ' contains 0
            Assert.IsFalse(Regex.IsMatch("COMM-2026-0001-K1X2", pattern)) ' contains 1
            Assert.IsFalse(Regex.IsMatch("COMM-2026-0001-KIX2", pattern)) ' contains I
            Assert.IsFalse(Regex.IsMatch("COMM-2026-0001-KLX2", pattern)) ' contains L
            Assert.IsFalse(Regex.IsMatch("COMM-2026-0001-KOX2", pattern)) ' contains O
        End Sub

        <TestMethod>
        Public Sub PortalSubmission_ModelProperties_AssignAndRetrieveCorrectly()
            Dim nowUtc = DateTime.UtcNow
            Dim subItem As New PortalSubmission With {
                .ControlNumber = "COMM-2026-0001-K9X2",
                .Category = "COMM",
                .DocumentTitle = "Formal Letter of Intent",
                .RequesterName = "Amina S. Macacua",
                .RequesterEmail = "amina.macacua@example.gov.ph",
                .RequesterPhone = "09171234567",
                .CreatedAt = nowUtc
            }

            Assert.AreEqual("COMM-2026-0001-K9X2", subItem.ControlNumber)
            Assert.AreEqual("COMM", subItem.Category)
            Assert.AreEqual("Formal Letter of Intent", subItem.DocumentTitle)
            Assert.AreEqual("Amina S. Macacua", subItem.RequesterName)
            Assert.AreEqual("amina.macacua@example.gov.ph", subItem.RequesterEmail)
            Assert.AreEqual("09171234567", subItem.RequesterPhone)
            Assert.AreEqual(nowUtc, subItem.CreatedAt)
        End Sub

        <TestMethod>
        Public Async Function FakePortalBridge_EnqueuesAndAcknowledgesCleanly() As Task
            Dim fake As New FakePortalBridge()
            fake.Queue.Add(New PortalSubmission With {
                .ControlNumber = "LEG-2026-0001-AB23",
                .Category = "LEG",
                .DocumentTitle = "Bill Submission",
                .RequesterName = "Atty. Omar",
                .RequesterEmail = "omar@example.com",
                .RequesterPhone = "09181234567",
                .CreatedAt = DateTime.UtcNow
            })

            Dim items = Await fake.FetchExternalQueueAsync()
            Assert.AreEqual(1, items.Count)

            Dim ack = Await fake.AcknowledgeImportAsync("LEG-2026-0001-AB23")
            Assert.IsTrue(ack)
            Assert.AreEqual(1, fake.AcknowledgedNumbers.Count)

            Dim itemsAfter = Await fake.FetchExternalQueueAsync()
            Assert.AreEqual(0, itemsAfter.Count, "Acknowledged item must be removed from queue.")

            Dim push = Await fake.PushPublicStatusAsync("LEG-2026-0001-AB23", "Under Review")
            Assert.IsTrue(push)
            Assert.AreEqual("Under Review", fake.StatusPushes("LEG-2026-0001-AB23"))
        End Function

        <TestMethod>
        Public Async Function DocumentService_PushPortalStatusSafe_PushesMappedStatusToBridge() As Task
            Dim fake As New FakePortalBridge()
            DocumentService.PushPortalStatusSafe("LEG-2026-0001-AB23", "APPROVED", fake)
            Await Task.Delay(150)
            Assert.IsTrue(fake.StatusPushes.ContainsKey("LEG-2026-0001-AB23"))
            Assert.AreEqual("Approved", fake.StatusPushes("LEG-2026-0001-AB23"))

            DocumentService.PushPortalStatusSafe("LEG-2026-0001-AB23", "RELEASED", fake)
            Await Task.Delay(150)
            Assert.AreEqual("Ready for Release", fake.StatusPushes("LEG-2026-0001-AB23"))

            DocumentService.PushPortalStatusSafe("LEG-2026-0001-AB23", "FOR_REVISION", fake)
            Await Task.Delay(150)
            Assert.AreEqual("For Processing", fake.StatusPushes("LEG-2026-0001-AB23"))

            DocumentService.PushPortalStatusSafe("LEG-2026-0001-AB23", "FOR_REVIEW", fake)
            Await Task.Delay(150)
            Assert.AreEqual("Under Review", fake.StatusPushes("LEG-2026-0001-AB23"))
        End Function

        <TestMethod>
        Public Sub DocumentService_PushPortalStatusSafe_WhenEmptyControlNumber_NoOpWithoutError()
            Dim fake As New FakePortalBridge()
            DocumentService.PushPortalStatusSafe("", "APPROVED", fake)
            DocumentService.PushPortalStatusSafe(Nothing, "APPROVED", fake)
            Assert.AreEqual(0, fake.StatusPushes.Count)
        End Sub
    End Class
End Namespace
