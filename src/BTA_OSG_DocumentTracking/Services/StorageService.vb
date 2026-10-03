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

        Public Function MoveDocument(docId As Integer, storageLocationId As Integer, movedByUserId As Integer, reason As String, Optional transaction As Microsoft.Data.SqlClient.SqlTransaction = Nothing) As DocumentMovement
            Dim movement As New DocumentMovement With {
                .DocumentID = docId,
                .StorageLocationID = storageLocationId,
                .MovedByUserID = movedByUserId,
                .MovedAtUTC = DateTime.UtcNow,
                .MovementReason = reason
            }

            If transaction IsNot Nothing Then
                ' The caller owns the transaction (MoveStorage wraps the landmark lookup and
                ' this movement together), so the audit row joins it: an audit failure after
                ' commit must not turn an already-committed move into a cache fallback that
                ' replays a duplicate movement.
                Dim newId As Integer = _storageRepo.InsertMovement(movement, transaction)
                movement.MovementID = newId
                _docRepo.UpdateStorageLocation(docId, storageLocationId, transaction)
                If _auditService IsNot Nothing Then
                    _auditService.LogEvent("STORAGE_LOCATION_CHANGED", "Storage", docId.ToString(), Nothing, Nothing, Nothing, True, Nothing, transaction)
                End If
            Else
                Using conn = _docRepo.ConnectionFactory.CreateConnection()
                    Using trans = conn.BeginTransaction()
                        Try
                            Dim newId As Integer = _storageRepo.InsertMovement(movement, trans)
                            movement.MovementID = newId
                            _docRepo.UpdateStorageLocation(docId, storageLocationId, trans)
                            If _auditService IsNot Nothing Then
                                _auditService.LogEvent("STORAGE_LOCATION_CHANGED", "Storage", docId.ToString(), Nothing, Nothing, Nothing, True, Nothing, trans)
                            End If
                            trans.Commit()
                        Catch
                            trans.Rollback()
                            Throw
                        End Try
                    End Using
                End Using
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
