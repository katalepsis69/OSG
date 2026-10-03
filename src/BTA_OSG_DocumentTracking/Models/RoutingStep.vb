Option Explicit On
Option Strict On

Imports System

Namespace BTA_OSG
    Public Class RoutingStep
        Public Property StepNumber As Integer
        Public Property StageCode As String
        Public Property StageName As String
        Public Property ResponsibleOffice As String
        Public Property ActionRequired As String
        Public Property StepStatus As String = "UPCOMING"
        Public Property CompletedAtUTC As DateTime?
        Public Property CompletedBy As String
        Public Property Remarks As String
    End Class
End Namespace
