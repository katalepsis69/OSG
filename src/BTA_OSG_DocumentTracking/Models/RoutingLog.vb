Option Explicit On
Option Strict On

Imports System

Namespace BTA_OSG
    Public Class RoutingLog
        Public Property RoutingLogID As Integer
        Public Property DocumentID As Integer
        Public Property FromStatusID As Integer?
        ' Nullable: a replayed log whose action is not a status code must not be coerced
        ' onto RECEIVED, and tbl_RoutingLogs.ToStatusID already accepts NULL.
        Public Property ToStatusID As Integer?
        Public Property FromOffice As String
        Public Property ToOffice As String
        Public Property RoutingRemarks As String
        Public Property RoutedByUserID As Integer
        Public Property RoutedAtUTC As DateTime
    End Class
End Namespace
