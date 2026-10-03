Option Explicit On
Option Strict On

Namespace BTA_OSG
    Public Class DocumentType
        Public Property DocumentTypeID As Integer
        Public Property TypeCode As String
        Public Property TypeName As String
        Public Property Prefix As String
        Public Property IsActive As Boolean
        Public Property SortOrder As Integer

        Public Shared Function IsValidCategory(code As String) As Boolean
            If String.IsNullOrWhiteSpace(code) Then Return False
            Dim c = code.Trim().ToUpperInvariant()
            Return c = "REG_COMM" OrElse c = "LEG" OrElse c = "FIN" OrElse c = "TRAVEL"
        End Function
    End Class
End Namespace
