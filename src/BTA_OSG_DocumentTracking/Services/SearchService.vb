Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic

Namespace BTA_OSG
    Public Class SearchService
        Private ReadOnly _docRepo As DocumentRepository

        Public Sub New(docRepo As DocumentRepository)
            _docRepo = docRepo
        End Sub

        Public Function Search(sessionContext As SessionContext, titleLike As String, typeId As Integer?, statusId As Integer?, originLike As String, destLike As String, storageId As Integer?, dateFrom As DateTime?, dateTo As DateTime?, pageSize As Integer, pageNumber As Integer) As List(Of Document)
            Dim hasViewAll As Boolean = sessionContext IsNot Nothing AndAlso sessionContext.Permissions IsNot Nothing AndAlso sessionContext.Permissions.Contains(RbacPolicy.DOCUMENT_VIEW_ALL)
            Dim userId As Integer = If(sessionContext IsNot Nothing AndAlso sessionContext.User IsNot Nothing, sessionContext.User.UserID, 0)
            Return _docRepo.SearchDocuments(hasViewAll, userId, titleLike, typeId, statusId, originLike, destLike, storageId, dateFrom, dateTo, pageSize, pageNumber)
        End Function
    End Class
End Namespace
