Imports System
Imports System.Collections.Generic

Namespace BTA_OSG
    Public Class UserService
        Private ReadOnly _userRepo As Object
        Private ReadOnly _auditService As Object

        Public Sub New(userRepo As Object, auditService As Object)
            _userRepo = userRepo
            _auditService = auditService
        End Sub

        Public Function GetAllUsers() As List(Of Object)
            Return _userRepo.GetAll()
        End Function

        Public Function CreateUser(user As Object, createdBy As Integer) As Integer
            Dim newId = _userRepo.Insert(user)
            _auditService.LogEvent("USER_CREATED", "User", newId, Nothing, Nothing, Nothing, True, Nothing)
            Return newId
        End Function

        Public Sub UpdateUser(user As Object, modifiedBy As Integer)
            _userRepo.Update(user)
            _auditService.LogEvent("USER_UPDATED", "User", user.UserID, Nothing, Nothing, Nothing, True, Nothing)
        End Sub

        Public Sub DisableUser(userId As Integer, disabledBy As Integer)
            _userRepo.Disable(userId)
            _auditService.LogEvent("USER_DISABLED", "User", userId, Nothing, Nothing, Nothing, True, Nothing)
        End Sub
    End Class
End Namespace
