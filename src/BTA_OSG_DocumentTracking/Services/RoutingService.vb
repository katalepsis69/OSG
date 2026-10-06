Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic

Imports System.Threading.Tasks

Namespace BTA_OSG
    ''' <summary>
    ''' Thrown when the server-side from-status guard refuses a transition. The
    ''' coordinator's offline fallback must rethrow this instead of caching the move:
    ''' replaying a transition the server deliberately rejected would file phantom
    ''' custody and audit rows for an action that never happened.
    ''' </summary>
    Public Class GuardRefusedException
        Inherits InvalidOperationException

        Public Sub New(message As String)
            MyBase.New(message)
        End Sub
    End Class

    Public Class RoutingService
        Private ReadOnly _routingRepo As RoutingRepository
        Private ReadOnly _docRepo As DocumentRepository
        Private ReadOnly _refRepo As ReferenceDataRepository
        Private ReadOnly _auditService As AuditService
        Private ReadOnly _portalBridge As IPortalBridge

        Public Sub New(routingRepo As RoutingRepository, docRepo As DocumentRepository, refRepo As ReferenceDataRepository, auditService As AuditService, Optional portalBridge As IPortalBridge = Nothing)
            _routingRepo = routingRepo
            _docRepo = docRepo
            _refRepo = refRepo
            _auditService = auditService
            _portalBridge = If(portalBridge, New PortalBridge(auditService:=auditService))
        End Sub

        Public Function ResubmitDocument(docId As Integer, resubmittedByUserId As Integer, notes As String, Optional expectedFromStatusId As Integer = 0) As Boolean
            Dim forReviewStatus = _refRepo.GetStatusByCode("FOR_REVIEW")
            Dim statusId As Integer = If(forReviewStatus IsNot Nothing, forReviewStatus.StatusID, 2)
            Dim doc = _docRepo.GetById(docId)
            Dim fromSec As String = If(doc IsNot Nothing AndAlso Not String.IsNullOrEmpty(doc.AssignedSection), doc.AssignedSection, "Section")

            Using conn = _docRepo.ConnectionFactory.CreateConnection()
                Using trans = conn.BeginTransaction()
                    Try
                        ApplyGuardedState(trans, docId, expectedFromStatusId, statusId, "Secretary-General", Nothing, "Resubmitted by " & fromSec & ": " & notes, resubmittedByUserId)
                        Dim routingLog As New RoutingLog With {
                            .DocumentID = docId,
                            .FromStatusID = If(doc IsNot Nothing, doc.StatusID, statusId),
                            .ToStatusID = statusId,
                            .FromOffice = fromSec,
                            .ToOffice = "Office of the Secretary-General",
                            .RoutingRemarks = "Resubmitted: " & notes,
                            .RoutedByUserID = resubmittedByUserId,
                            .RoutedAtUTC = DateTime.UtcNow
                        }
                        _routingRepo.Insert(routingLog, trans)
                        ' The audit row joins the transaction: an audit failure after a bare
                        ' commit used to reach the coordinator's offline fallback and re-apply
                        ' the transition locally.
                        If _auditService IsNot Nothing Then
                            _auditService.LogEvent("DOCUMENT_RESUBMITTED", "Document", docId.ToString(),
                                                   If(doc IsNot Nothing, doc.DocCode, Nothing),
                                                   StatusJson(doc), """Status"":""FOR_REVIEW""", True, Nothing, trans)
                        End If
                        trans.Commit()
                    Catch
                        trans.Rollback()
                        Throw
                    End Try
                End Using
            End Using

            NotifyPortalStatus(docId, "FOR_REVIEW")
            Return True
        End Function

        Public Function ApproveDocument(docId As Integer, approvedByUserId As Integer, notes As String, Optional expectedFromStatusId As Integer = 0) As Boolean
            Dim approvedStatus = _refRepo.GetStatusByCode("APPROVED")
            Dim statusId As Integer = If(approvedStatus IsNot Nothing, approvedStatus.StatusID, 6)
            Dim doc = _docRepo.GetById(docId)

            Using conn = _docRepo.ConnectionFactory.CreateConnection()
                Using trans = conn.BeginTransaction()
                    Try
                        ApplyGuardedState(trans, docId, expectedFromStatusId, statusId, "Records Section", Nothing, "Approved by Secretary-General: " & notes, approvedByUserId)
                        Dim routingLog As New RoutingLog With {
                            .DocumentID = docId,
                            .FromStatusID = If(doc IsNot Nothing, doc.StatusID, statusId),
                            .ToStatusID = statusId,
                            .FromOffice = "Office of the Secretary-General",
                            .ToOffice = "Records Section",
                            .RoutingRemarks = "Approved: " & notes,
                            .RoutedByUserID = approvedByUserId,
                            .RoutedAtUTC = DateTime.UtcNow
                        }
                        _routingRepo.Insert(routingLog, trans)
                        If _auditService IsNot Nothing Then
                            _auditService.LogEvent("DOCUMENT_APPROVED", "Document", docId.ToString(),
                                                   If(doc IsNot Nothing, doc.DocCode, Nothing),
                                                   StatusJson(doc), """Status"":""APPROVED""", True, Nothing, trans)
                        End If
                        trans.Commit()
                    Catch
                        trans.Rollback()
                        Throw
                    End Try
                End Using
            End Using

            NotifyPortalStatus(docId, "APPROVED")
            Return True
        End Function

        Public Function ReleaseDocument(docId As Integer, releasedByUserId As Integer, notes As String, Optional expectedFromStatusId As Integer = 0) As Boolean
            Dim releasedStatus = _refRepo.GetStatusByCode("RELEASED")
            Dim statusId As Integer = If(releasedStatus IsNot Nothing, releasedStatus.StatusID, 7)
            Dim doc = _docRepo.GetById(docId)

            Using conn = _docRepo.ConnectionFactory.CreateConnection()
                Using trans = conn.BeginTransaction()
                    Try
                        ApplyGuardedState(trans, docId, expectedFromStatusId, statusId, "Archived / Released", Nothing, "Released to destination office: " & notes, releasedByUserId)
                        Dim routingLog As New RoutingLog With {
                            .DocumentID = docId,
                            .FromStatusID = If(doc IsNot Nothing, doc.StatusID, statusId),
                            .ToStatusID = statusId,
                            .FromOffice = "Records Section",
                            .ToOffice = If(doc IsNot Nothing AndAlso Not String.IsNullOrEmpty(doc.DestinationOffice), doc.DestinationOffice, "Destination Office"),
                            .RoutingRemarks = "Released: " & notes,
                            .RoutedByUserID = releasedByUserId,
                            .RoutedAtUTC = DateTime.UtcNow
                        }
                        _routingRepo.Insert(routingLog, trans)
                        If _auditService IsNot Nothing Then
                            _auditService.LogEvent("DOCUMENT_RELEASED", "Document", docId.ToString(),
                                                   If(doc IsNot Nothing, doc.DocCode, Nothing),
                                                   StatusJson(doc), """Status"":""RELEASED""", True, Nothing, trans)
                        End If
                        trans.Commit()
                    Catch
                        trans.Rollback()
                        Throw
                    End Try
                End Using
            End Using

            NotifyPortalStatus(docId, "RELEASED")
            Return True
        End Function

        ''' <summary>
        ''' Before/after bag for a transition audit row: the status (and the desk it sat at)
        ''' the document is leaving. AGENTS.md Rule 4 wants JSON snapshots, not just the fact
        ''' that something happened.
        ''' </summary>
        Private Shared Function StatusJson(doc As Document) As String
            Dim code As String = If(doc IsNot Nothing, DesktopDataCoordinator.StatusCodeFor(doc.StatusID), "")
            Dim section As String = If(doc IsNot Nothing, doc.AssignedSection, "")
            Return "{""Status"":""" & AuditService.JsonText(code) & """,""Section"":""" & AuditService.JsonText(section) & """}"
        End Function

        ''' <summary>
        ''' One guarded workflow update for the connected transitions: when the caller states
        ''' the status its form was showing, the write lands only if the server row still
        ''' carries it, so a document another seat already moved cannot be approved, released,
        ''' or resubmitted a second time. Zero means no guard (callers that have no snapshot).
        ''' </summary>
        Private Sub ApplyGuardedState(trans As Microsoft.Data.SqlClient.SqlTransaction, docId As Integer, expectedFromStatusId As Integer, statusId As Integer, assignedSection As String, punchlist As String, lastAction As String, modifiedBy As Integer)
            If expectedFromStatusId > 0 Then
                Dim applied = _docRepo.UpdateWorkflowStateFromStatus(docId, expectedFromStatusId, statusId, assignedSection, punchlist, lastAction, modifiedBy, transaction:=trans)
                If applied = 0 Then Throw New GuardRefusedException("The document's state changed on the server. Refresh the document and try again.")
            Else
                _docRepo.UpdateWorkflowState(docId, statusId, assignedSection, punchlist, lastAction, modifiedBy, transaction:=trans)
            End If
        End Sub

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

        Public Function GetRoutingHistory(docId As Integer) As List(Of RoutingLog)
            Return _routingRepo.GetByDocumentId(docId)
        End Function
    End Class
End Namespace
