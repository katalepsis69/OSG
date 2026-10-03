Option Explicit On
Option Strict On

Imports System

Namespace BTA_OSG
    Public Class DocumentMovement
        Public Property MovementID As Integer
        Public Property DocumentID As Integer
        Public Property StorageLocationID As Integer
        Public Property MovedByUserID As Integer
        Public Property MovedAtUTC As DateTime
        Public Property MovementReason As String
    End Class
End Namespace
