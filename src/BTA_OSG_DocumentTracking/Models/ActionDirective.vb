Imports System

Namespace BTA_OSG
    Public Class ActionDirective
        Public Property DirectiveID As Integer
        Public Property DocumentID As Integer
        Public Property DirectiveTypeID As Integer
        Public Property DirectiveText As String
        Public Property IssuedByUserID As Integer
        Public Property IssuedAtUTC As DateTime
        Public Property IsActive As Boolean
        Public Property SupersededByDirectiveID As Integer?
        Public Property Remarks As String
    End Class
End Namespace
