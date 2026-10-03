Option Explicit On
Option Strict On

Imports System.Collections.Generic
Imports System.Threading.Tasks

Namespace BTA_OSG
    ''' <summary>
    ''' Contract for interacting with the external public intake portal bridge API.
    ''' </summary>
    Public Interface IPortalBridge
        ''' <summary>
        ''' Retrieves the list of unimported public submissions from the external portal queue.
        ''' </summary>
        Function FetchExternalQueueAsync() As Task(Of List(Of PortalSubmission))

        ''' <summary>
        ''' Notifies the external portal that a submission was imported into the internal registry.
        ''' </summary>
        Function AcknowledgeImportAsync(controlNumber As String) As Task(Of Boolean)

        ''' <summary>
        ''' Notifies the external portal that a batch of submissions was imported into the internal registry.
        ''' </summary>
        Function AcknowledgeBatchImportAsync(controlNumbers As IEnumerable(Of String)) As Task(Of Boolean)

        ''' <summary>
        ''' Pushes a public status update and logs a public milestone on the external portal.
        ''' </summary>
        Function PushPublicStatusAsync(controlNumber As String, publicStatus As String) As Task(Of Boolean)
    End Interface
End Namespace
