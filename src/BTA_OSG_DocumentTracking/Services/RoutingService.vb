Imports System
Imports System.Collections.Generic

Namespace BTA_OSG
    Public Class RoutingService
        Private ReadOnly _routingRepo As Object
        Private ReadOnly _docRepo As Object
        Private ReadOnly _refRepo As Object
        Private ReadOnly _auditService As Object

        Public Sub New(routingRepo As Object, docRepo As Object, refRepo As Object, auditService As Object)
            _routingRepo = routingRepo
            _docRepo = docRepo
            _refRepo = refRepo
            _auditService = auditService
        End Sub

        Public Function RouteDocument(docId As Integer, fromStatusId As Integer, toStatusCode As String, fromOffice As Integer, toOffice As Integer, remarks As String, routedByUserId As Integer) As Object
            Dim toStatus = _refRepo.GetStatusByCode(toStatusCode)
            
            Dim routingLog = Nothing ' New RoutingLog
            _routingRepo.Insert(routingLog)
            _docRepo.UpdateStatus(docId, toStatus.StatusID)
            
            _auditService.LogEvent("ROUTING_LOG_ADDED", "RoutingLog", docId, Nothing, Nothing, Nothing, True, Nothing)
            Return routingLog
        End Function

        Public Function GetRoutingHistory(docId As Integer) As List(Of Object)
            Return _routingRepo.GetByDocId(docId)
        End Function
    End Class
End Namespace
