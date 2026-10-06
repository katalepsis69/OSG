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
        Private ReadOnly _routingRepo As RoutingRepository

        Public Sub New(directiveRepo As DirectiveRepository, docRepo As DocumentRepository, refRepo As ReferenceDataRepository, auditService As AuditService, Optional portalBridge As IPortalBridge = Nothing, Optional routingRepo As RoutingRepository = Nothing)
            _directiveRepo = directiveRepo
            _docRepo = docRepo
            _refRepo = refRepo
            _auditService = auditService
            _portalBridge = If(portalBridge, New PortalBridge(auditService:=auditService))
            _routingRepo = routingRepo
        End Sub

        Public Function IssueDirective(docId As Integer, directiveTypeId As Integer, directiveText As String, remarks As String, issuedByUserId As Integer, Optional assignedTo As String = "") As ActionDirective
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

            ' The status side effect must not reopen a finished document (for example the
            ' "Approved & Archived" directive on a RELEASED row), and the refusal has to
            ' reach the operator rather than fall back to the offline cache.
            Dim doc = _docRepo.GetById(docId)
            If targetStatusId > 0 AndAlso doc IsNot Nothing Then
                Dim currentCode = DesktopDataCoordinator.StatusCodeFor(doc.StatusID)
                Dim targetCode = DesktopDataCoordinator.StatusCodeFor(targetStatusId)
                If DocumentStatus.IsTerminalStatus(currentCode) AndAlso
                   Not String.Equals(currentCode, targetCode, StringComparison.OrdinalIgnoreCase) Then
                    Throw New GuardRefusedException("This document is already " & DocumentStatus.DisplayName(currentCode) & " and can no longer receive this directive.")
                End If
            End If

            ' The assignee combo carries the staff name; resolve it once, before the
            ' transaction, so the assignment row joins the same fact as the directive.
            Dim assigneeId As Integer? = DesktopDataCoordinator.ResolveUserId(assignedTo)

            ' Directive insert, status side effect, custody log, assignment, and audit row
            ' are one fact: a mid-way failure must not leave a status change the directive
            ' record does not support.
            Using conn = _docRepo.ConnectionFactory.CreateConnection()
                Using tx = conn.BeginTransaction()
                    Try
                        directive.DirectiveID = _directiveRepo.Insert(directive, tx)
                        If targetStatusId > 0 Then
                            _docRepo.UpdateStatus(docId, targetStatusId, tx)
                            If _routingRepo IsNot Nothing Then
                                _routingRepo.Insert(New RoutingLog With {
                                    .DocumentID = docId,
                                    .FromStatusID = If(doc IsNot Nothing, doc.StatusID, CType(Nothing, Integer?)),
                                    .ToStatusID = targetStatusId,
                                    .FromOffice = If(doc IsNot Nothing, doc.AssignedSection, ""),
                                    .ToOffice = If(doc IsNot Nothing, doc.AssignedSection, ""),
                                    .RoutingRemarks = "SG Directive: " & directiveText,
                                    .RoutedByUserID = issuedByUserId,
                                    .RoutedAtUTC = DateTime.UtcNow
                                }, tx)
                            End If
                        End If
                        If assigneeId.HasValue Then
                            _docRepo.AddAssignment(New DocumentAssignment With {
                                .DocumentID = docId,
                                .AssignedUserID = assigneeId.Value,
                                .AssignedByUserID = issuedByUserId,
                                .AssignedAtUTC = DateTime.UtcNow,
                                .Remarks = "Assigned via SG directive"
                            }, tx)
                        End If
                        If _auditService IsNot Nothing Then
                            _auditService.LogEvent("DIRECTIVE_ADDED", "Directive", docId.ToString(),
                                                   If(doc IsNot Nothing, doc.DocCode, Nothing),
                                                   StatusJson(doc),
                                                   "{""Directive"":""" & AuditService.JsonText(directiveText) & """,""AssignedTo"":""" & AuditService.JsonText(assignedTo) & """}",
                                                   True, Nothing, tx)
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

        ''' <summary>
        ''' Before/after bag for the audit row: the status (and desk) the document is
        ''' leaving. AGENTS.md Rule 4 wants JSON snapshots, not just the fact.
        ''' </summary>
        Private Shared Function StatusJson(doc As Document) As String
            Dim code As String = If(doc IsNot Nothing, DesktopDataCoordinator.StatusCodeFor(doc.StatusID), "")
            Dim section As String = If(doc IsNot Nothing, doc.AssignedSection, "")
            Return "{""Status"":""" & AuditService.JsonText(code) & """,""Section"":""" & AuditService.JsonText(section) & """}"
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

            ' A revision order must not drag a finished document back into the workflow.
            If doc IsNot Nothing AndAlso DocumentStatus.IsTerminalStatus(DesktopDataCoordinator.StatusCodeFor(doc.StatusID)) Then
                Throw New GuardRefusedException("This document is already " & DocumentStatus.DisplayName(DesktopDataCoordinator.StatusCodeFor(doc.StatusID)) & " and can no longer be returned for revision.")
            End If

            Using conn = _docRepo.ConnectionFactory.CreateConnection()
                Using tx = conn.BeginTransaction()
                    Try
                        directive.DirectiveID = _directiveRepo.Insert(directive, tx)
                        _docRepo.UpdateWorkflowState(docId, statusId, targetSec, combinedPunchlist, "Sec Gen requested revision: " & punchlistNotes, issuedByUserId, transaction:=tx)
                        If _auditService IsNot Nothing Then
                            _auditService.LogEvent("REVISION_REQUESTED", "Directive", docId.ToString(),
                                                   If(doc IsNot Nothing, doc.DocCode, Nothing),
                                                   StatusJson(doc),
                                                   "{""Status"":""FOR_REVISION"",""Section"":""" & AuditService.JsonText(targetSec) & """}",
                                                   True, Nothing, tx)
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
