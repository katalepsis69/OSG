Imports System
Imports System.Collections.Generic

Namespace BTA_OSG
    Public Class SessionManager
        Public Shared Property CurrentSession As SessionContext

        Public Shared Sub Login(session As SessionContext)
            CurrentSession = session
        End Sub

        Public Shared Sub Logout(auditRepo As Object)
            If CurrentSession IsNot Nothing Then
                ' Log logout event
                CurrentSession = Nothing
            End If
        End Sub

        Public Shared Sub SetCurrentSession(session As SessionContext)
            Login(session)
        End Sub

        Public Shared Sub EndSession()
            Logout(Nothing)
        End Sub

        Public Shared Function CheckTimeout(settings As AppSettings, auditRepo As Object) As Boolean
            If CurrentSession Is Nothing Then Return True
            
            Dim timeSinceLastActivity As TimeSpan = DateTime.UtcNow - CurrentSession.LastActivityUTC
            If timeSinceLastActivity.TotalMinutes > settings.SessionSettings.TimeoutMinutes Then
                ' Log timeout
                CurrentSession = Nothing
                Return True
            End If
            
            Return False
        End Function

        Public Shared Sub Touch()
            If CurrentSession IsNot Nothing Then
                CurrentSession.LastActivityUTC = DateTime.UtcNow
            End If
        End Sub

        Public Shared Function HasPermission(code As String) As Boolean
            If CurrentSession Is Nothing Then Return False
            Return CurrentSession.Permissions.Contains(code)
        End Function

        Public Shared ReadOnly Property IsLoggedIn As Boolean
            Get
                Return CurrentSession IsNot Nothing
            End Get
        End Property
    End Class
End Namespace
