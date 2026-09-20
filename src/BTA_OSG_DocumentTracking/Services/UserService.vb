Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic

Namespace BTA_OSG
    Public Class UserService
        Private ReadOnly _userRepo As UserRepository
        Private ReadOnly _auditService As AuditService

        Public Sub New(userRepo As UserRepository, auditService As AuditService)
            _userRepo = userRepo
            _auditService = auditService
        End Sub

        Public Function GetAllUsers() As List(Of User)
            Return _userRepo.GetAll()
        End Function

        Public Function CreateUser(user As User, createdBy As Integer) As Integer
            Dim newId = _userRepo.Insert(user, createdBy)
            If _auditService IsNot Nothing Then
                _auditService.LogEvent("USER_CREATED", "User", newId.ToString(), Nothing, Nothing, Nothing, True, Nothing)
            End If
            Return newId
        End Function

        Public Sub UpdateUser(user As User, modifiedBy As Integer)
            _userRepo.Update(user, modifiedBy)
            If _auditService IsNot Nothing Then
                _auditService.LogEvent("USER_UPDATED", "User", user.UserID.ToString(), Nothing, Nothing, Nothing, True, Nothing)
            End If
        End Sub

        Public Sub DisableUser(userId As Integer, disabledBy As Integer)
            _userRepo.Disable(userId)
            If _auditService IsNot Nothing Then
                _auditService.LogEvent("USER_DISABLED", "User", userId.ToString(), Nothing, Nothing, Nothing, True, Nothing)
            End If
        End Sub
    End Class
End Namespace
