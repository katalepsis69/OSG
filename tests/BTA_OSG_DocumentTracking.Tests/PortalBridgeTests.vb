Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
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
        Public Async Function DocumentService_PushPortalStatusSafe_PushesMappedStatusToBridge() As Task
            Dim fake As New FakePortalBridge()
            DocumentService.PushPortalStatusSafe("LEG-2026-0001-AB23", "APPROVED", fake)
            Await WaitForStatusAsync(fake, "LEG-2026-0001-AB23")
            Assert.IsTrue(fake.StatusPushes.ContainsKey("LEG-2026-0001-AB23"))
            Assert.AreEqual("Approved", fake.StatusPushes("LEG-2026-0001-AB23"))

            DocumentService.PushPortalStatusSafe("LEG-2026-0001-AB23", "RELEASED", fake)
            Await WaitForStatusAsync(fake, "LEG-2026-0001-AB23", "Ready for Release")
            Assert.AreEqual("Ready for Release", fake.StatusPushes("LEG-2026-0001-AB23"))

            DocumentService.PushPortalStatusSafe("LEG-2026-0001-AB23", "FOR_REVISION", fake)
            Await WaitForStatusAsync(fake, "LEG-2026-0001-AB23", "For Processing")
            Assert.AreEqual("For Processing", fake.StatusPushes("LEG-2026-0001-AB23"))

            DocumentService.PushPortalStatusSafe("LEG-2026-0001-AB23", "FOR_REVIEW", fake)
            Await WaitForStatusAsync(fake, "LEG-2026-0001-AB23", "Under Review")
            Assert.AreEqual("Under Review", fake.StatusPushes("LEG-2026-0001-AB23"))
        End Function

        ''' <summary>
        ''' The push is fire-and-forget; a single fixed sleep loses the race on a loaded
        ''' runner, so poll until the mapped status lands (or the budget expires).
        ''' </summary>
        Private Shared Async Function WaitForStatusAsync(fake As FakePortalBridge, controlNumber As String, Optional expectedStatus As String = Nothing, Optional timeoutMs As Integer = 5000) As Task
            Dim deadline = DateTime.UtcNow.AddMilliseconds(timeoutMs)
            While DateTime.UtcNow < deadline
                Dim current As String = Nothing
                If fake.StatusPushes.TryGetValue(controlNumber, current) AndAlso (expectedStatus Is Nothing OrElse current = expectedStatus) Then Return
                Await Task.Delay(25)
            End While
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
