Option Explicit On
Option Strict On

Imports System

Namespace BTA_OSG
    ''' <summary>
    ''' Public portal submission record pending ingestion into the OSG internal document registry.
    ''' </summary>
    Public Class PortalSubmission
        Public Property ControlNumber As String
        Public Property Category As String
        Public Property DocumentTitle As String
        Public Property RequesterName As String
        Public Property RequesterEmail As String
        Public Property RequesterPhone As String
        Public Property RequesterGender As String = ""
        Public Property CreatedAt As DateTime
    End Class
End Namespace
