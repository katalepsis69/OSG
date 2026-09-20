Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic

Namespace BTA_OSG
    Public Class DocumentService
        Private ReadOnly _docRepo As DocumentRepository
        Private ReadOnly _seqRepo As SequenceRepository
        Private ReadOnly _refRepo As ReferenceDataRepository
        Private ReadOnly _auditService As AuditService

        Public Sub New(docRepo As DocumentRepository, seqRepo As SequenceRepository, refRepo As ReferenceDataRepository, auditService As AuditService)
            _docRepo = docRepo
            _seqRepo = seqRepo
            _refRepo = refRepo
            _auditService = auditService
        End Sub

        Public Function RegisterDocument(title As String, typeCode As String, originOffice As String, destOffice As String, receivedDate As DateTime, googleDriveUrl As String, remarks As String, registeredByUserId As Integer) As Document
            If String.IsNullOrWhiteSpace(title) Then Throw New ArgumentException("Title is required.")
            Dim prefix As String = If(String.IsNullOrWhiteSpace(typeCode), "DOC", typeCode.Trim().ToUpperInvariant())
            Dim code As String = _seqRepo.GetNextDocCode(prefix, prefix, CShort(DateTime.Now.Year))

            Dim docType = _refRepo.GetDocumentTypeByCode(prefix)
            Dim typeId As Integer = If(docType IsNot Nothing, docType.DocumentTypeID, 1)

            Dim doc As New Document With {
                .DocCode = code,
                .Title = title,
                .DocumentTypeID = typeId,
                .OriginOffice = originOffice,
                .DestinationOffice = destOffice,
                .StatusID = 1,
                .ReceivedDate = receivedDate,
                .GoogleDriveUrl = googleDriveUrl,
                .Remarks = remarks,
                .RegisteredByUserID = registeredByUserId,
                .RegisteredAtUTC = DateTime.UtcNow
            }
            Dim newId As Integer = _docRepo.Insert(doc)
            doc.DocumentID = newId

            If _auditService IsNot Nothing Then
                _auditService.LogEvent("DOCUMENT_CREATED", "Document", doc.DocumentID.ToString(), code, Nothing, Nothing, True, Nothing)
            End If
            Return doc
        End Function

        Public Function RegisterDocument(title As String, typeCode As String, originOffice As Integer, destOffice As Integer, receivedDate As DateTime, googleDriveUrl As String, remarks As String, registeredByUserId As Integer) As Document
            Return RegisterDocument(title, typeCode, originOffice.ToString(), destOffice.ToString(), receivedDate, googleDriveUrl, remarks, registeredByUserId)
        End Function

        Public Sub UpdateDocument(doc As Document, modifiedBy As Integer)
            _docRepo.Update(doc, modifiedBy)
            If _auditService IsNot Nothing Then
                _auditService.LogEvent("DOCUMENT_UPDATED", "Document", doc.DocumentID.ToString(), doc.DocCode, Nothing, Nothing, True, Nothing)
            End If
        End Sub

        Public Sub SoftDeleteDocument(docId As Integer, deletedBy As Integer, reason As String)
            _docRepo.SoftDelete(docId, deletedBy, reason)
            If _auditService IsNot Nothing Then
                _auditService.LogEvent("DOCUMENT_DELETED", "Document", docId.ToString(), Nothing, Nothing, Nothing, True, reason)
            End If
        End Sub

        Public Function GetDocument(docId As Integer) As Document
            Return _docRepo.GetById(docId)
        End Function

        Public Function SearchDocuments(titleLike As String, typeId As Integer?, statusId As Integer?, originLike As String, destLike As String, storageId As Integer?, dateFrom As DateTime?, dateTo As DateTime?, pageSize As Integer, pageNumber As Integer) As List(Of Document)
            Return _docRepo.GetByFilter(titleLike, typeId, statusId, originLike, destLike, storageId, dateFrom, dateTo, pageSize, pageNumber)
        End Function
    End Class
End Namespace
