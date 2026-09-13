Imports System
Imports System.Collections.Generic

Namespace BTA_OSG
    Public Class StorageService
        Private ReadOnly _storageRepo As Object
        Private ReadOnly _docRepo As Object
        Private ReadOnly _auditService As Object

        Public Sub New(storageRepo As Object, docRepo As Object, auditService As Object)
            _storageRepo = storageRepo
            _docRepo = docRepo
            _auditService = auditService
        End Sub

        Public Function MoveDocument(docId As Integer, storageLocationId As Integer, movedByUserId As Integer, reason As String) As Object
            Dim movement = Nothing ' New DocumentMovement
            _storageRepo.Insert(movement)
            _docRepo.UpdateStorageLocation(docId, storageLocationId)
            
            _auditService.LogEvent("STORAGE_LOCATION_CHANGED", "Storage", docId, Nothing, Nothing, Nothing, True, Nothing)
            Return movement
        End Function

        Public Function GetMovementHistory(docId As Integer) As List(Of Object)
            Return _storageRepo.GetMovementsByDocId(docId)
        End Function

        Public Function GetAllLocations() As List(Of Object)
            Return _storageRepo.GetAllLocations()
        End Function
    End Class
End Namespace
