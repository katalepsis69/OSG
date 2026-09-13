Imports System
Imports System.Collections.Generic

Namespace BTA_OSG
    Public Class SessionContext
        Public Property SessionID As Guid
        Public Property User As User
        Public Property Roles As List(Of Role)
        Public Property Permissions As HashSet(Of String)
        Public Property LoginAtUTC As DateTime
        Public Property LastActivityUTC As DateTime

        Public Sub New()
            SessionID = Guid.NewGuid()
            Roles = New List(Of Role)()
            Permissions = New HashSet(Of String)()
            LoginAtUTC = DateTime.UtcNow
            LastActivityUTC = DateTime.UtcNow
        End Sub

        Public Function HasPermission(code As String) As Boolean
            If Permissions Is Nothing Then Return False
            Return Permissions.Contains(code)
        End Function

        Public Function IsExpired(timeoutMinutes As Integer) As Boolean
            Return (DateTime.UtcNow - LastActivityUTC).TotalMinutes > timeoutMinutes
        End Function

        Public Sub Touch()
            LastActivityUTC = DateTime.UtcNow
        End Sub
    End Class
End Namespace
