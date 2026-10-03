Option Explicit On
Option Strict On

Imports System

Namespace BTA_OSG
    Public Class RfidCard
        Public Property RfidCardID As Integer
        Public Property UserID As Integer
        Public Property CardPublicID As String
        Public Property CardLabel As String
        Public Property IsActive As Boolean
        Public Property IssuedAtUTC As DateTime
        Public Property RevokedAtUTC As DateTime?
        Public Property RevokedByUserID As Integer?
        Public Property RevocationReason As String
    End Class
End Namespace
