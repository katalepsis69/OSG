Imports System
Imports System.Collections.Generic

Namespace BTA_OSG
    Public Class SearchService
        Private ReadOnly _docRepo As Object

        Public Sub New(docRepo As Object)
            _docRepo = docRepo
        End Sub

        Public Function Search(sessionContext As Object, titleLike As String, typeId As Integer?, statusId As Integer?, originLike As String, destLike As String, storageId As Integer?, dateFrom As DateTime?, dateTo As DateTime?, pageSize As Integer, pageNumber As Integer) As List(Of Object)
            Dim hasViewAll As Boolean = sessionContext.Permissions.Contains(RbacPolicy.DOCUMENT_VIEW_ALL)
            Return _docRepo.SearchDocuments(hasViewAll, sessionContext.UserID, titleLike, typeId, statusId, originLike, destLike, storageId, dateFrom, dateTo, pageSize, pageNumber)
        End Function
    End Class
End Namespace
