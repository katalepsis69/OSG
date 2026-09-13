Imports System

Namespace BTA_OSG
    Public Class RfidFailedAttempt
        Public Property FailedAttemptID As Long
        Public Property CardPublicID As String
        Public Property UserID As Integer?
        Public Property AttemptAtUTC As DateTime
        Public Property MachineName As String
        Public Property Reason As String
        Public Property Success As Boolean
    End Class
End Namespace
