Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic

Namespace BTA_OSG
    Public Class DirectiveService
        Private ReadOnly _directiveRepo As DirectiveRepository
        Private ReadOnly _docRepo As DocumentRepository
        Private ReadOnly _refRepo As ReferenceDataRepository
        Private ReadOnly _auditService As AuditService

        Public Sub New(directiveRepo As DirectiveRepository, docRepo As DocumentRepository, refRepo As ReferenceDataRepository, auditService As AuditService)
            _directiveRepo = directiveRepo
            _docRepo = docRepo
            _refRepo = refRepo
            _auditService = auditService
        End Sub

        Public Function IssueDirective(docId As Integer, directiveTypeId As Integer, directiveText As String, remarks As String, issuedByUserId As Integer) As ActionDirective
            Dim directive As New ActionDirective With {
                .DocumentID = docId,
                .DirectiveTypeID = directiveTypeId,
                .DirectiveText = directiveText,
                .IssuedByUserID = issuedByUserId,
                .IssuedAtUTC = DateTime.UtcNow,
                .IsActive = True,
                .Remarks = remarks
            }
            Dim newId As Integer = _directiveRepo.Insert(directive)
            directive.DirectiveID = newId

            Dim directiveTypes = _refRepo.GetDirectiveTypes()
            Dim targetStatusId As Integer = 0
            For Each dt In directiveTypes
                If dt.DirectiveTypeID = directiveTypeId AndAlso dt.ResultStatusID.HasValue Then
                    targetStatusId = dt.ResultStatusID.Value
                    Exit For
                End If
            Next

            If targetStatusId > 0 Then
                _docRepo.UpdateStatus(docId, targetStatusId)
            End If

            If _auditService IsNot Nothing Then
                _auditService.LogEvent("DIRECTIVE_ADDED", "Directive", docId.ToString(), Nothing, Nothing, Nothing, True, Nothing)
            End If
            Return directive
        End Function

        Public Function GetDirectives(docId As Integer) As List(Of ActionDirective)
            Return _directiveRepo.GetByDocId(docId)
        End Function
    End Class
End Namespace
