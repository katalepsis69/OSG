Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic

Imports System.Threading.Tasks

Namespace BTA_OSG
    Public Class DocumentService
        Private ReadOnly _docRepo As DocumentRepository
        Private ReadOnly _seqRepo As SequenceRepository
        Private ReadOnly _refRepo As ReferenceDataRepository
        Private ReadOnly _auditService As AuditService
        Private ReadOnly _portalBridge As IPortalBridge
        Private ReadOnly _routingRepo As RoutingRepository
        Private ReadOnly _storageRepo As StorageRepository

        Public Sub New(docRepo As DocumentRepository, seqRepo As SequenceRepository, refRepo As ReferenceDataRepository, auditService As AuditService, Optional portalBridge As IPortalBridge = Nothing, Optional routingRepo As RoutingRepository = Nothing, Optional storageRepo As StorageRepository = Nothing)
            _docRepo = docRepo
            _seqRepo = seqRepo
            _refRepo = refRepo
            _auditService = auditService
            _portalBridge = If(portalBridge, New PortalBridge(auditService:=auditService))
            _routingRepo = routingRepo
            _storageRepo = storageRepo
        End Sub

        Public Shared Function GetDefaultSectionForCategory(categoryCode As String) As String
            If String.IsNullOrWhiteSpace(categoryCode) Then Return "Records Section"
            Select Case categoryCode.Trim().ToUpperInvariant()
                Case "REG_COMM", "COMMUNICATION", "REGULAR COMMUNICATION" : Return "Secretariat"
                Case "LEG", "LEGISLATIVE" : Return "Legislative Section"
                Case "FIN", "FINANCE" : Return "Finance Section"
                Case "TRAVEL", "TRAVEL ORDER", "TO" : Return "Travel Section"
                Case Else : Return "Records Section"
            End Select
        End Function

        Public Function RegisterDocument(title As String, typeCode As String, originOffice As String, destOffice As String, receivedDate As DateTime, googleDriveUrl As String, remarks As String, registeredByUserId As Integer) As Document
            Return RegisterDocumentWithWorkflow(title, typeCode, "INCOMING", originOffice, destOffice, Nothing, googleDriveUrl, remarks, registeredByUserId, receivedDate)
        End Function

        ''' <summary>
        ''' Creates the document row. When the caller supplies a transaction, the sequence
        ''' reservation, the insert, and the audit entry all join it, so a caller wrapping
        ''' the whole registration workflow can roll back every write together.
        ''' preferredDocCode (the code minted offline) is reused when free and re-minted on
        ''' a unique-key collision with another workstation's sequence.
        ''' </summary>
        Public Function RegisterDocumentWithWorkflow(title As String, typeCode As String, flowDirection As String, originOffice As String, destOffice As String, targetDeadline As DateTime?, googleDriveUrl As String, remarks As String, registeredByUserId As Integer, Optional receivedDate As DateTime? = Nothing, Optional transaction As Microsoft.Data.SqlClient.SqlTransaction = Nothing, Optional externalControlNumber As String = "", Optional preferredDocCode As String = "") As Document
            If String.IsNullOrWhiteSpace(title) Then Throw New ArgumentException("Title is required.")
            Dim prefix As String = If(String.IsNullOrWhiteSpace(typeCode), "DOC", typeCode.Trim().ToUpperInvariant())
            If prefix = "REG_COMM" Then prefix = "COMM"
            If prefix = "TRAVEL" Then prefix = "TO"

            ' The type resolves by the caller's TypeCode, which is the seeded vocabulary
            ' (REG_COMM/LEG/FIN/TRAVEL); only the sequence key uses the shorter mint prefix.
            Dim docType = _refRepo.GetDocumentTypeByCode(If(String.IsNullOrWhiteSpace(typeCode), "DOC", typeCode.Trim().ToUpperInvariant()))
            Dim typeId As Integer = If(docType IsNot Nothing, docType.DocumentTypeID, 1)

            Dim assignedSec As String = GetDefaultSectionForCategory(typeCode)
            Dim receivedStatus = _refRepo.GetStatusByCode("RECEIVED")
            Dim statusId As Integer = If(receivedStatus IsNot Nothing, receivedStatus.StatusID, 1)

            Dim doc As New Document With {
                .Title = title,
                .DocumentTypeID = typeId,
                .FlowDirection = If(String.IsNullOrWhiteSpace(flowDirection), "INCOMING", flowDirection.Trim().ToUpperInvariant()),
                .OriginOffice = originOffice,
                .DestinationOffice = destOffice,
                .AssignedSection = assignedSec,
                .StatusID = statusId,
                .ReceivedDate = If(receivedDate.HasValue, receivedDate.Value, DateTime.Now),
                .TargetDeadlineUTC = targetDeadline,
                .GoogleDriveUrl = googleDriveUrl,
                .Remarks = remarks,
                .LastActionTaken = "Registered by Records Section and routed to " & assignedSec,
                .RegisteredByUserID = registeredByUserId,
                .RegisteredAtUTC = DateTime.UtcNow,
                .ExternalControlNumber = externalControlNumber
            }

            If Not String.IsNullOrWhiteSpace(preferredDocCode) Then
                doc.DocCode = preferredDocCode.Trim().ToUpperInvariant()
                Try
                    Dim preferredId As Integer = _docRepo.Insert(doc, transaction)
                    doc.DocumentID = preferredId
                    ' The offline number skipped the counter, so the counter has to be told.
                    _seqRepo.RaiseSequenceTo(doc.DocCode, prefix, transaction)
                    LogCreated(doc, registeredByUserId, transaction)
                    Return doc
                Catch dup As Microsoft.Data.SqlClient.SqlException When IsDuplicateKey(dup)
                    ' Another workstation's sequence already minted this number offline;
                    ' fall through to the shared sequence so the insert still lands.
                End Try
            End If

            Dim code As String = _seqRepo.GetNextDocCode(prefix, prefix, CShort(DateTime.Now.Year), transaction)
            doc.DocCode = code
            Dim newId As Integer = _docRepo.Insert(doc, transaction)
            doc.DocumentID = newId
            LogCreated(doc, registeredByUserId, transaction)
            Return doc
        End Function

        Private Sub LogCreated(doc As Document, registeredByUserId As Integer, transaction As Microsoft.Data.SqlClient.SqlTransaction)
            If _auditService IsNot Nothing Then
                _auditService.LogEvent("DOCUMENT_CREATED", "Document", doc.DocumentID.ToString(), doc.DocCode, Nothing, Nothing, True, Nothing, transaction)
            End If
        End Sub

        Private Shared Function IsDuplicateKey(ex As Microsoft.Data.SqlClient.SqlException) As Boolean
            For i As Integer = 0 To ex.Errors.Count - 1
                If ex.Errors(i).Number = 2601 OrElse ex.Errors(i).Number = 2627 Then Return True
            Next
            Return False
        End Function

        Public Function RegisterDocument(doc As Document) As Integer
            Dim newId As Integer = _docRepo.Insert(doc)
            doc.DocumentID = newId
            If _auditService IsNot Nothing Then
                _auditService.LogEvent("DOCUMENT_CREATED", "Document", doc.DocumentID.ToString(), doc.DocCode, Nothing, Nothing, True, Nothing)
            End If
            Return newId
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

        Public Shared Function MapToPublicStatus(statusCode As String) As String
            If String.IsNullOrWhiteSpace(statusCode) Then Return "Received"
            Select Case statusCode.Trim().ToUpperInvariant()
                Case "RECEIVED"
                    Return "Received"
                Case "FOR_REVIEW", "PENDING_REVIEW"
                    Return "Under Review"
                Case "FOR_REVISION", "REVISION_REQUESTED", "PROCESSING"
                    Return "For Processing"
                Case "APPROVED"
                    Return "Approved"
                Case "RELEASED", "FILED"
                    Return "Ready for Release"
                Case "REJECTED", "CANCELLED"
                    Return "Rejected"
                Case Else
                    Return "Received"
            End Select
        End Function

        Public Function RegisterImportedExternalDocument(submission As PortalSubmission, registeredByUserId As Integer) As Document
            If submission Is Nothing Then Throw New ArgumentNullException(NameOf(submission))
            If String.IsNullOrWhiteSpace(submission.ControlNumber) Then Throw New ArgumentException("Control number is required.")

            ' Deduplication guard: check if already imported
            Dim existing = _docRepo.GetByExternalControlNumber(submission.ControlNumber)
            If existing IsNot Nothing Then
                Return existing
            End If

            Dim category = If(String.IsNullOrWhiteSpace(submission.Category), "COMM", submission.Category.Trim().ToUpperInvariant())
            Dim prefix As String = category
            If prefix = "REG_COMM" Then prefix = "COMM"
            If prefix = "TRAVEL" Then prefix = "TO"

            ' The seeded TypeCodes are REG_COMM/LEG/FIN/TRAVEL, so the type resolves by the
            ' unmapped category and never by the mint prefix (COMM/TO are sequence keys only).
            Dim typeCode = category
            If typeCode = "COMM" Then typeCode = "REG_COMM"
            If typeCode = "TO" Then typeCode = "TRAVEL"
            Dim docType = _refRepo.GetDocumentTypeByCode(typeCode)
            Dim typeId As Integer = If(docType IsNot Nothing, docType.DocumentTypeID, 1)

            Dim assignedSec As String = GetDefaultSectionForCategory(category)
            Dim receivedStatus = _refRepo.GetStatusByCode("RECEIVED")
            Dim statusId As Integer = If(receivedStatus IsNot Nothing, receivedStatus.StatusID, 1)

            Dim origin = If(Not String.IsNullOrWhiteSpace(submission.RequesterName), submission.RequesterName & " (" & submission.RequesterEmail & ")", "Public Portal Intake")

            Dim doc As New Document With {
                .Title = submission.DocumentTitle,
                .DocumentTypeID = typeId,
                .FlowDirection = "INCOMING",
                .OriginOffice = origin,
                .DestinationOffice = "Office of the Secretary-General",
                .AssignedSection = assignedSec,
                .StatusID = statusId,
                .ReceivedDate = submission.CreatedAt,
                .Remarks = "Imported from Public Intake Portal. Submitter: " & submission.RequesterName & " (" & submission.RequesterEmail & ", " & submission.RequesterPhone & "). Gender: " & If(String.IsNullOrWhiteSpace(submission.RequesterGender), "Prefer not to say", submission.RequesterGender) & ". External CN: " & submission.ControlNumber,
                .LastActionTaken = "Imported from Public Portal and routed to " & assignedSec,
                .RegisteredByUserID = registeredByUserId,
                .RegisteredAtUTC = DateTime.UtcNow,
                .ExternalControlNumber = submission.ControlNumber
            }

            ' The document row, its workflow state, the routing log, and the initial storage
            ' movement are one fact: they land whole in one transaction or not at all, so a
            ' mirror row can never replay them into a second registration.
            Using conn = _docRepo.ConnectionFactory.CreateConnection()
                Using tx = conn.BeginTransaction()
                    Try
                        doc.DocCode = _seqRepo.GetNextDocCode(prefix, prefix, CShort(DateTime.Now.Year), tx)
                        Dim newId As Integer = _docRepo.Insert(doc, tx)
                        doc.DocumentID = newId

                        _docRepo.UpdateWorkflowState(newId, statusId, assignedSec, Nothing, doc.LastActionTaken, registeredByUserId, destinationOffice:=doc.DestinationOffice, transaction:=tx)

                        If _routingRepo IsNot Nothing Then
                            _routingRepo.Insert(New RoutingLog With {
                                .DocumentID = newId,
                                .FromStatusID = Nothing,
                                .ToStatusID = statusId,
                                .FromOffice = "Public Portal",
                                .ToOffice = doc.DestinationOffice,
                                .RoutedByUserID = registeredByUserId,
                                .RoutedAtUTC = DateTime.UtcNow,
                                .RoutingRemarks = "Imported from Public Intake Portal."
                            }, tx)
                        End If

                        If _storageRepo IsNot Nothing Then
                            Dim locationId = _storageRepo.GetOrCreateByKey("UNFILED", "", "", tx)
                            _storageRepo.InsertMovement(New DocumentMovement With {
                                .DocumentID = newId,
                                .StorageLocationID = locationId,
                                .MovedByUserID = registeredByUserId,
                                .MovedAtUTC = DateTime.UtcNow,
                                .MovementReason = "Initial storage assignment (awaiting physical filing)."
                            }, tx)
                            _docRepo.UpdateStorageLocation(newId, locationId, tx)
                        End If

                        If _auditService IsNot Nothing Then
                            _auditService.LogEvent("PORTAL_DOCUMENT_IMPORTED", "Document", newId.ToString(), doc.DocCode, Nothing, Nothing, True, "External CN: " & submission.ControlNumber, tx)
                        End If

                        tx.Commit()
                    Catch
                        tx.Rollback()
                        Throw
                    End Try
                End Using
            End Using

            If _portalBridge IsNot Nothing Then
                Task.Run(Async Function()
                             Await _portalBridge.AcknowledgeImportAsync(submission.ControlNumber).ConfigureAwait(False)
                         End Function)
            End If

            Return doc
        End Function

        Public Async Function SyncPortalStatusAsync(docId As Integer, statusCode As String) As Task(Of Boolean)
            If _portalBridge Is Nothing Then Return True
            Try
                Dim extCn As String = ""
                Try
                    Dim doc = _docRepo.GetById(docId)
                    If doc IsNot Nothing Then extCn = doc.ExternalControlNumber
                Catch
                    ' Offline / fallback
                End Try

                If String.IsNullOrWhiteSpace(extCn) Then
                    Dim rows = EmbeddedDB.DataSet.Tables("Documents").Select("DocumentID = " & docId)
                    If rows.Length > 0 AndAlso rows(0).Table.Columns.Contains("ExternalControlNumber") AndAlso Not IsDBNull(rows(0)("ExternalControlNumber")) Then
                        extCn = rows(0)("ExternalControlNumber").ToString()
                    End If
                End If

                If Not String.IsNullOrWhiteSpace(extCn) Then
                    Dim publicStatus = MapToPublicStatus(statusCode)
                    Return Await _portalBridge.PushPublicStatusAsync(extCn, publicStatus).ConfigureAwait(False)
                End If
            Catch ex As Exception
                ' Resilient isolation
            End Try
            Return True
        End Function

        ''' <summary>
        ''' Safe, non-blocking fire-and-forget public portal status push conforming to the isolation invariant:
        ''' No portal network failure or outage may ever crash or block internal workflows.
        ''' </summary>
        Public Shared Sub PushPortalStatusSafe(externalControlNumber As String, statusCode As String, Optional portalBridge As IPortalBridge = Nothing)
            If String.IsNullOrWhiteSpace(externalControlNumber) Then Return
            Dim bridge = If(portalBridge, AppStartup.PortalBridgeClient)
            If bridge Is Nothing Then Return
            Dim publicStatus = MapToPublicStatus(statusCode)
            Task.Run(Async Function()
                         Try
                             Await bridge.PushPublicStatusAsync(externalControlNumber, publicStatus).ConfigureAwait(False)
                         Catch ex As Exception
                             ' Resilient isolation invariant
                         End Try
                     End Function)
        End Sub
    End Class
End Namespace
