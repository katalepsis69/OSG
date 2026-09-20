Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic

Namespace BTA_OSG
    Public Class StorageService
        Private ReadOnly _storageRepo As StorageRepository
        Private ReadOnly _docRepo As DocumentRepository
        Private ReadOnly _auditService As AuditService

        Public Sub New(storageRepo As StorageRepository, docRepo As DocumentRepository, auditService As AuditService)
            _storageRepo = storageRepo
            _docRepo = docRepo
            _auditService = auditService
        End Sub

        Public Function MoveDocument(docId As Integer, storageLocationId As Integer, movedByUserId As Integer, reason As String) As DocumentMovement
            Dim movement As New DocumentMovement With {
                .DocumentID = docId,
                .StorageLocationID = storageLocationId,
                .MovedByUserID = movedByUserId,
                .MovedAtUTC = DateTime.UtcNow,
                .MovementReason = reason
            }
            Dim newId As Integer = _storageRepo.InsertMovement(movement)
            movement.MovementID = newId

            _docRepo.UpdateStorageLocation(docId, storageLocationId)

            If _auditService IsNot Nothing Then
                _auditService.LogEvent("STORAGE_LOCATION_CHANGED", "Storage", docId.ToString(), Nothing, Nothing, Nothing, True, Nothing)
            End If
            Return movement
        End Function

        Public Function GetMovementHistory(docId As Integer) As List(Of DocumentMovement)
            Return _storageRepo.GetMovements(docId)
        End Function

        Public Function GetAllLocations() As List(Of StorageLocation)
            Return _storageRepo.GetAll()
        End Function
    End Class
End Namespace
