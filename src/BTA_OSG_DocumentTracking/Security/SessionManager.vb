Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic

Namespace BTA_OSG
    Public Class SessionManager
        Public Shared Property CurrentSession As SessionContext

        Public Shared Sub Login(session As SessionContext)
            CurrentSession = session
        End Sub

        Public Shared Sub Logout(Optional auditRepo As Object = Nothing)
            If CurrentSession IsNot Nothing Then
                CurrentSession = Nothing
            End If
        End Sub

        Public Shared Sub Touch()
            If CurrentSession IsNot Nothing Then
                CurrentSession.LastActivityUTC = DateTime.UtcNow
            End If
        End Sub

        ' Permission alias map ensuring backward-compatible permission code checks
        Private Shared ReadOnly _aliasMap As New Dictionary(Of String, String()) From {
            {RbacPolicy.DOCUMENT_VIEW_ALL, {RbacPolicy.DOC_VIEW, "DASHBOARD_VIEW"}},
            {RbacPolicy.DOCUMENT_CREATE, {RbacPolicy.DOC_CREATE}},
            {RbacPolicy.DOCUMENT_UPDATE, {RbacPolicy.DOC_EDIT}},
            {RbacPolicy.DIRECTIVE_ISSUE, {RbacPolicy.DOC_DIRECTIVE}},
            {RbacPolicy.ROUTING_CREATE, {RbacPolicy.DOC_ROUTE}},
            {RbacPolicy.STORAGE_UPDATE, {RbacPolicy.STORAGE_MANAGE}},
            {RbacPolicy.ROLE_MANAGE, {RbacPolicy.SYS_CONFIG}},
            {RbacPolicy.REPORT_EXPORT, {RbacPolicy.REP_VIEW}}
        }
        Public Shared Function HasPermission(code As String) As Boolean
            If CurrentSession Is Nothing OrElse CurrentSession.Permissions Is Nothing Then Return False
            If CurrentSession.Permissions.Contains(code) Then Return True
            Dim aliases As String() = Nothing
            If _aliasMap.TryGetValue(code, aliases) Then
                For Each a In aliases
                    If CurrentSession.Permissions.Contains(a) Then Return True
                Next
            End If
            Return False
        End Function

        Public Shared ReadOnly Property IsLoggedIn As Boolean
            Get
                Return CurrentSession IsNot Nothing
            End Get
        End Property
    End Class
End Namespace
