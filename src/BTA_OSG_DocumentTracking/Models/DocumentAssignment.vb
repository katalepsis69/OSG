Option Explicit On
Option Strict On

Imports System

Namespace BTA_OSG
    Public Class DocumentAssignment
        Public Property AssignmentID As Integer
        Public Property DocumentID As Integer
        Public Property AssignedUserID As Integer?
        Public Property AssignedOffice As String
        Public Property AssignedRoleID As Integer?
        Public Property AssignedByUserID As Integer
        Public Property AssignedAtUTC As DateTime
        Public Property IsActive As Boolean
        Public Property Remarks As String
    End Class
End Namespace
