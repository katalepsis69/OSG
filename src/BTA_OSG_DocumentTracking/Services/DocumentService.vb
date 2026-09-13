Imports System
Imports System.Collections.Generic

Namespace BTA_OSG
    Public Class DocumentService
        Private ReadOnly _docRepo As Object
        Private ReadOnly _seqRepo As Object
        Private ReadOnly _refRepo As Object
        Private ReadOnly _auditService As Object

        Public Sub New(docRepo As Object, seqRepo As Object, refRepo As Object, auditService As Object)
            _docRepo = docRepo
            _seqRepo = seqRepo
            _refRepo = refRepo
            _auditService = auditService
        End Sub

        Public Function RegisterDocument(title As String, typeCode As String, originOffice As Integer, destOffice As Integer, receivedDate As DateTime, googleDriveUrl As String, remarks As String, registeredByUserId As Integer) As Object
            If String.IsNullOrWhiteSpace(title) Then Throw New ArgumentException("Title is required.")
            ' Mock get type
            Dim prefix As String = "DOC" 
            Dim code As String = _seqRepo.GetNextDocCode(prefix, DateTime.Now.Year)
            
            Dim doc As Object = Nothing ' New Document with details
            _docRepo.Insert(doc)
            _auditService.LogEvent("DOCUMENT_CREATED", "Document", doc.DocumentID, code, Nothing, Nothing, True, Nothing)
            Return doc
        End Function

        Public Sub UpdateDocument(doc As Object, modifiedBy As Integer)
            _docRepo.Update(doc)
            _auditService.LogEvent("DOCUMENT_UPDATED", "Document", doc.DocumentID, doc.DocumentCode, Nothing, Nothing, True, Nothing)
        End Sub

        Public Sub SoftDeleteDocument(docId As Integer, deletedBy As Integer, reason As String)
            _docRepo.SoftDelete(docId, deletedBy, reason)
            _auditService.LogEvent("DOCUMENT_DELETED", "Document", docId, Nothing, Nothing, Nothing, True, reason)
        End Sub

        Public Function GetDocument(docId As Integer) As Object
            Return _docRepo.GetById(docId)
        End Function

        Public Function SearchDocuments(filterParams As Object) As List(Of Object)
            Return _docRepo.GetVisibleToUser(filterParams)
        End Function
    End Class
End Namespace
