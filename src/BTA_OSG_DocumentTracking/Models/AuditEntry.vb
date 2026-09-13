Imports System

Namespace BTA_OSG
    Public Class AuditEntry
        Public Property AuditID As Long
        Public Property EventAtUTC As DateTime
        Public Property SessionID As Guid?
        Public Property UserID As Integer?
        Public Property UsernameSnapshot As String
        Public Property FullNameSnapshot As String
        Public Property RoleSnapshot As String
        Public Property ActionType As String
        Public Property EntityType As String
        Public Property EntityID As String
        Public Property DocumentCode As String
        Public Property OldValuesJson As String
        Public Property NewValuesJson As String
        Public Property MachineName As String
        Public Property ClientInfo As String
        Public Property Success As Boolean
        Public Property FailureReason As String
        Public Property CorrelationId As Guid?
        Public Property CardPublicIDMasked As String
        Public Property ApplicationVersion As String
    End Class
End Namespace
