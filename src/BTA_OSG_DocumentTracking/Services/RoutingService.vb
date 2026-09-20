Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic

Namespace BTA_OSG
    Public Class RoutingService
        Private ReadOnly _routingRepo As RoutingRepository
        Private ReadOnly _docRepo As DocumentRepository
        Private ReadOnly _refRepo As ReferenceDataRepository
        Private ReadOnly _auditService As AuditService

        Public Sub New(routingRepo As RoutingRepository, docRepo As DocumentRepository, refRepo As ReferenceDataRepository, auditService As AuditService)
            _routingRepo = routingRepo
            _docRepo = docRepo
            _refRepo = refRepo
            _auditService = auditService
        End Sub

        Public Function RouteDocument(docId As Integer, fromStatusId As Integer, toStatusCode As String, fromOffice As String, toOffice As String, remarks As String, routedByUserId As Integer) As RoutingLog
            Dim toStatus = _refRepo.GetStatusByCode(toStatusCode)
            Dim toStatusId As Integer = If(toStatus IsNot Nothing, toStatus.StatusID, fromStatusId)

            Dim routingLog As New RoutingLog With {
                .DocumentID = docId,
                .FromStatusID = fromStatusId,
                .ToStatusID = toStatusId,
                .FromOffice = fromOffice,
                .ToOffice = toOffice,
                .RoutingRemarks = remarks,
                .RoutedByUserID = routedByUserId,
                .RoutedAtUTC = DateTime.UtcNow
            }
            Dim newId As Integer = _routingRepo.Insert(routingLog)
            routingLog.RoutingLogID = newId

            _docRepo.UpdateStatus(docId, toStatusId)

            If _auditService IsNot Nothing Then
                _auditService.LogEvent("ROUTING_LOG_ADDED", "RoutingLog", docId.ToString(), Nothing, Nothing, Nothing, True, Nothing)
            End If
            Return routingLog
        End Function

        Public Function GetRoutingHistory(docId As Integer) As List(Of RoutingLog)
            Return _routingRepo.GetByDocumentId(docId)
        End Function
    End Class
End Namespace
