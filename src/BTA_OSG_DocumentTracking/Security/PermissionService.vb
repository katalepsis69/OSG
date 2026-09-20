Imports System
Imports System.Collections.Generic

Namespace BTA_OSG
    Public Class PermissionService
        Private ReadOnly _userRepo As UserRepository

        Public Sub New(userRepo As UserRepository)
            _userRepo = userRepo
        End Sub

        Public Function GetEffectivePermissions(userId As Integer) As HashSet(Of String)
            Return _userRepo.GetUserPermissions(userId)
        End Function

        Public Function HasPermission(userId As Integer, permissionCode As String) As Boolean
            Dim perms = GetEffectivePermissions(userId)
            Return perms.Contains(permissionCode)
        End Function
    End Class
End Namespace
