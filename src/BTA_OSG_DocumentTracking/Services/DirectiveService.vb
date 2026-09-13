Imports System
Imports System.Collections.Generic

Namespace BTA_OSG
    Public Class DirectiveService
        Private ReadOnly _directiveRepo As Object
        Private ReadOnly _docRepo As Object
        Private ReadOnly _refRepo As Object
        Private ReadOnly _auditService As Object

        Public Sub New(directiveRepo As Object, docRepo As Object, refRepo As Object, auditService As Object)
            _directiveRepo = directiveRepo
            _docRepo = docRepo
            _refRepo = refRepo
            _auditService = auditService
        End Sub

        Public Function IssueDirective(docId As Integer, directiveTypeId As Integer, directiveText As String, remarks As String, issuedByUserId As Integer) As Object
            Dim directive = Nothing ' New Directive
            _directiveRepo.Insert(directive)

            Dim typeStatus = _refRepo.GetResultStatusForDirective(directiveTypeId)
            If typeStatus IsNot Nothing Then
                _docRepo.UpdateStatus(docId, typeStatus.StatusID)
            Else
                _docRepo.UpdateStatus(docId, "DIRECTIVE_ISSUED")
            End If

            _auditService.LogEvent("DIRECTIVE_ADDED", "Directive", docId, Nothing, Nothing, Nothing, True, Nothing)
            Return directive
        End Function

        Public Function GetDirectives(docId As Integer) As List(Of Object)
            Return _directiveRepo.GetByDocId(docId)
        End Function
    End Class
End Namespace
