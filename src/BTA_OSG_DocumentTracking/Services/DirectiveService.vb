Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic

Imports System.Threading.Tasks

Namespace BTA_OSG
    Public Class DirectiveService
        Private ReadOnly _directiveRepo As DirectiveRepository
        Private ReadOnly _docRepo As DocumentRepository
        Private ReadOnly _refRepo As ReferenceDataRepository
        Private ReadOnly _auditService As AuditService
        Private ReadOnly _portalBridge As IPortalBridge

        Public Sub New(directiveRepo As DirectiveRepository, docRepo As DocumentRepository, refRepo As ReferenceDataRepository, auditService As AuditService, Optional portalBridge As IPortalBridge = Nothing)
            _directiveRepo = directiveRepo
            _docRepo = docRepo
            _refRepo = refRepo
            _auditService = auditService
            _portalBridge = If(portalBridge, New PortalBridge(auditService:=auditService))
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

            Dim targetStatusId As Integer = 0
            For Each dt In _refRepo.GetDirectiveTypes()
                If dt.DirectiveTypeID = directiveTypeId AndAlso dt.ResultStatusID.HasValue Then
                    targetStatusId = dt.ResultStatusID.Value
                    Exit For
                End If
            Next

            ' Directive insert, status side effect, and audit row are one fact: a mid-way
            ' failure must not leave a status change the directive record does not support.
            Using conn = _docRepo.ConnectionFactory.CreateConnection()
                Using tx = conn.BeginTransaction()
                    Try
                        directive.DirectiveID = _directiveRepo.Insert(directive, tx)
                        If targetStatusId > 0 Then
                            _docRepo.UpdateStatus(docId, targetStatusId, tx)
                        End If
                        If _auditService IsNot Nothing Then
                            _auditService.LogEvent("DIRECTIVE_ADDED", "Directive", docId.ToString(), Nothing, Nothing, Nothing, True, Nothing, tx)
                        End If
                        tx.Commit()
                    Catch
                        tx.Rollback()
                        Throw
                    End Try
                End Using
            End Using

            If targetStatusId > 0 Then
                Dim statuses = _refRepo.GetDocumentStatuses()
                For Each s As DocumentStatus In statuses
                    If s.StatusID = targetStatusId Then
                        NotifyPortalStatus(docId, s.StatusCode)
                        Exit For
                    End If
                Next
            End If
            Return directive
        End Function

        Public Function RequestRevision(docId As Integer, punchlistNotes As String, issuedByUserId As Integer, returnToSection As String) As ActionDirective
            Dim directiveTypeId As Integer = 1
            Dim directiveTypes = _refRepo.GetDirectiveTypes()
            For Each dt In directiveTypes
                If dt.DirectiveCode.IndexOf("REVIS", StringComparison.OrdinalIgnoreCase) >= 0 Then
                    directiveTypeId = dt.DirectiveTypeID
                    Exit For
                End If
            Next

            Dim directive As New ActionDirective With {
                .DocumentID = docId,
                .DirectiveTypeID = directiveTypeId,
                .DirectiveText = "Revision Requested: " & punchlistNotes,
                .IssuedByUserID = issuedByUserId,
                .IssuedAtUTC = DateTime.UtcNow,
                .IsActive = True,
                .Remarks = punchlistNotes
            }

            Dim forRevisionStatus = _refRepo.GetStatusByCode("FOR_REVISION")
            Dim statusId As Integer = If(forRevisionStatus IsNot Nothing, forRevisionStatus.StatusID, 5)

            Dim doc = _docRepo.GetById(docId)
            Dim targetSec As String = If(Not String.IsNullOrWhiteSpace(returnToSection), returnToSection, If(doc IsNot Nothing, doc.AssignedSection, "Records Section"))
            Dim combinedPunchlist As String = punchlistNotes
            If doc IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(doc.RevisionPunchlist) Then
                combinedPunchlist = doc.RevisionPunchlist & vbCrLf & "[" & DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm") & "] " & punchlistNotes
            End If

            Using conn = _docRepo.ConnectionFactory.CreateConnection()
                Using tx = conn.BeginTransaction()
                    Try
                        directive.DirectiveID = _directiveRepo.Insert(directive, tx)
                        _docRepo.UpdateWorkflowState(docId, statusId, targetSec, combinedPunchlist, "Sec Gen requested revision: " & punchlistNotes, issuedByUserId, transaction:=tx)
                        If _auditService IsNot Nothing Then
                            _auditService.LogEvent("REVISION_REQUESTED", "Directive", docId.ToString(), Nothing, Nothing, Nothing, True, Nothing, tx)
                        End If
                        tx.Commit()
                    Catch
                        tx.Rollback()
                        Throw
                    End Try
                End Using
            End Using

            NotifyPortalStatus(docId, "FOR_REVISION")
            Return directive
        End Function

        Private Sub NotifyPortalStatus(docId As Integer, statusCode As String)
            If _portalBridge Is Nothing Then Return
            Try
                Dim extCn As String = ""
                Try
                    Dim doc = _docRepo.GetById(docId)
                    If doc IsNot Nothing Then extCn = doc.ExternalControlNumber
                Catch
                    ' Offline fallback
                End Try

                If String.IsNullOrWhiteSpace(extCn) Then
                    Dim rows = EmbeddedDB.DataSet.Tables("Documents").Select("DocumentID = " & docId)
                    If rows.Length > 0 AndAlso rows(0).Table.Columns.Contains("ExternalControlNumber") AndAlso Not IsDBNull(rows(0)("ExternalControlNumber")) Then
                        extCn = rows(0)("ExternalControlNumber").ToString()
                    End If
                End If

                If Not String.IsNullOrWhiteSpace(extCn) Then
                    DocumentService.PushPortalStatusSafe(extCn, statusCode, _portalBridge)
                End If
            Catch ex As Exception
                ' Resilient isolation
            End Try
        End Sub

        Public Function GetDirectives(docId As Integer) As List(Of ActionDirective)
            Return _directiveRepo.GetByDocumentId(docId)
        End Function
    End Class
End Namespace
