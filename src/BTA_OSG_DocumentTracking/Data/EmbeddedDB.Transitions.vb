Option Explicit On
Option Strict On

Imports System
Imports System.Data

Namespace BTA_OSG
    ' Offline custody-transition verbs: directive/routing/movement/audit log writes and the
    ' RequestRevision, Resubmit, Approve and Release transitions. Split from EmbeddedDB.vb
    ' (same partial class, zero behavior change) so the store file stays under the 1k-line
    ' review threshold; schema setup and query helpers remain in the main file.
    Partial Public Class EmbeddedDB

        Public Shared Sub AddDirective(docId As Integer, directive As String, assignedTo As String, notes As String, staffName As String, Optional isOffline As Boolean = True, Optional issuedByUserId As Integer = 1, Optional directiveCode As String = "IMMEDIATE_ACTION")
            SyncLock _syncLock
                Dim dt = DataSet.Tables("Directives")
                Dim r = dt.NewRow()
                r("DocumentID") = docId
                r("SGDirective") = directive
                r("AssignedTo") = assignedTo
                r("Notes") = notes
                r("LogUser") = staffName
                r("Timestamp") = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                If dt.Columns.Contains("DirectiveCode") Then r("DirectiveCode") = directiveCode
                If dt.Columns.Contains("IssuedByUserID") Then r("IssuedByUserID") = issuedByUserId
                If dt.Columns.Contains("PendingSync") Then r("PendingSync") = isOffline
                dt.Rows.Add(r)

                Dim docRows = DataSet.Tables("Documents").Select("DocumentID = " & docId)
                If docRows.Length > 0 Then
                    ' Only directives with a real status outcome move the status column; the
                    ' directive text itself belongs in LastActionTaken. Free text in
                    ' CurrentStatus replayed to a RECEIVED fallback corrupted the status.
                    ' "Approved & Archived" maps to ARCHIVED to match the connected path,
                    ' where the APPROVE_ARCHIVE type's ResultStatusID is the archived status.
                    Dim targetCode As String = ""
                    If directive = "REVISION_REQUESTED" Then
                        targetCode = "FOR_REVISION"
                    ElseIf directive = "APPROVED" Then
                        targetCode = "APPROVED"
                    ElseIf directive = "Approved & Archived" Then
                        targetCode = "ARCHIVED"
                    End If
                    If targetCode.Length > 0 Then
                        Dim currentCode = docRows(0)("CurrentStatus").ToString().Trim().ToUpperInvariant()
                        If DocumentStatus.IsTerminalStatus(currentCode) AndAlso currentCode <> targetCode Then
                            Throw New InvalidOperationException("This document is already " & DocumentStatus.DisplayName(currentCode) & " and can no longer receive this directive.")
                        End If
                        docRows(0)("CurrentStatus") = targetCode
                    End If
                    docRows(0)("LastActionTaken") = "SG Directive: " & directive
                    If Not String.IsNullOrEmpty(assignedTo) Then docRows(0)("AssignedStaff") = assignedTo
                    If docRows(0).Table.Columns.Contains("ModifiedByUserID") Then docRows(0)("ModifiedByUserID") = issuedByUserId
                    If isOffline AndAlso docRows(0).Table.Columns.Contains("PendingSync") Then
                        docRows(0)("PendingSync") = True
                    End If
                End If
                MarkDirty()
            End SyncLock
        End Sub

        Public Shared Sub AddRoutingLog(docId As Integer, fromOffice As String, toOffice As String, routedBy As String, action As String, remarks As String, Optional isOffline As Boolean = True, Optional routedByUserId As Integer = 1)
            SyncLock _syncLock
                Dim dt = DataSet.Tables("RoutingLogs")
                Dim r = dt.NewRow()
                r("DocumentID") = docId
                r("FromOffice") = fromOffice
                r("ToOffice") = toOffice
                r("RoutedBy") = routedBy
                r("ActionTaken") = action
                r("Remarks") = remarks
                r("Timestamp") = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                If dt.Columns.Contains("RoutedByUserID") Then r("RoutedByUserID") = routedByUserId
                If dt.Columns.Contains("PendingSync") Then r("PendingSync") = isOffline
                dt.Rows.Add(r)

                Dim docRows = DataSet.Tables("Documents").Select("DocumentID = " & docId)
                If docRows.Length > 0 Then
                    docRows(0)("OriginatingOffice") = fromOffice
                    docRows(0)("DestinationOffice") = toOffice
                    ' Route targets are one fact with the custody log (the connected path moves
                    ' status and logs in one transaction), so the offline path moves status too.
                    ' The membership test reads the canonical list so the combo and this mover
                    ' can never drift apart. Internal markers (REGISTERED, REVISION_REQUESTED,
                    ' RESUBMITTED) are not status codes: their callers already set CurrentStatus
                    ' and must not be overridden here. AssignedSection also stays untouched:
                    ' ReleaseDocument sets "Archived / Released", which a destination office
                    ' must not overwrite.
                    Dim actionCode = If(action, "").Trim().ToUpperInvariant()
                    If Array.IndexOf(DocumentStatus.RouteTargetStatusCodes, actionCode) >= 0 Then
                        Dim currentCode = docRows(0)("CurrentStatus").ToString().Trim().ToUpperInvariant()
                        ' Re-logging the status the document already sits in is the internal
                        ' transitions re-recording their own target; moving a terminal
                        ' document anywhere else reopens a filed row and is refused.
                        If DocumentStatus.IsTerminalStatus(currentCode) AndAlso currentCode <> actionCode Then
                            Throw New InvalidOperationException("This document is already " & DocumentStatus.DisplayName(currentCode) & " and can no longer be routed.")
                        End If
                        docRows(0)("CurrentStatus") = actionCode
                    End If
                    If Not action.Equals("REVISION_REQUESTED", StringComparison.OrdinalIgnoreCase) Then
                        docRows(0)("LastActionTaken") = "Routed to " & toOffice & ": " & action
                    End If
                    If docRows(0).Table.Columns.Contains("ModifiedByUserID") Then docRows(0)("ModifiedByUserID") = routedByUserId
                    If isOffline AndAlso docRows(0).Table.Columns.Contains("PendingSync") Then
                        docRows(0)("PendingSync") = True
                    End If
                End If
                MarkDirty()
            End SyncLock
        End Sub

        Public Shared Sub AddMovementLog(docId As Integer, fromLoc As String, toLoc As String, movedBy As String, reason As String, Optional isOffline As Boolean = True, Optional movedByUserId As Integer = 1)
            SyncLock _syncLock
                Dim dt = DataSet.Tables("Movements")
                Dim r = dt.NewRow()
                r("DocumentID") = docId
                r("FromLocation") = fromLoc
                r("ToLocation") = toLoc
                r("MovedBy") = movedBy
                r("Reason") = reason
                r("Timestamp") = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                If dt.Columns.Contains("MovedByUserID") Then r("MovedByUserID") = movedByUserId
                If dt.Columns.Contains("PendingSync") Then r("PendingSync") = isOffline
                dt.Rows.Add(r)

                Dim parts = toLoc.Split(New Char() {"/"c, "|"c})
                Dim docRows = DataSet.Tables("Documents").Select("DocumentID = " & docId)
                If docRows.Length > 0 Then
                    If parts.Length >= 3 Then
                        docRows(0)("CabinetID") = parts(0).Trim()
                        docRows(0)("ShelfNo") = parts(1).Trim()
                        docRows(0)("BoxCode") = parts(2).Trim()
                    End If
                    docRows(0)("LastActionTaken") = "Moved storage to " & toLoc
                    If docRows(0).Table.Columns.Contains("ModifiedByUserID") Then docRows(0)("ModifiedByUserID") = movedByUserId
                    If isOffline AndAlso docRows(0).Table.Columns.Contains("PendingSync") Then
                        docRows(0)("PendingSync") = True
                    End If
                End If
                MarkDirty()
            End SyncLock
        End Sub

        Public Shared Sub LogAudit(user As String, action As String, Optional isOffline As Boolean = True, Optional userId As Integer = 1, Optional actionType As String = "")
            SyncLock _syncLock
                Dim dt = DataSet.Tables("AuditTrail")
                Dim r = dt.NewRow()
                r("UserName") = user
                r("ActionDescription") = action
                r("Timestamp") = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                If dt.Columns.Contains("ActionType") Then r("ActionType") = actionType
                If dt.Columns.Contains("UserID") Then r("UserID") = userId
                If dt.Columns.Contains("PendingSync") Then r("PendingSync") = isOffline
                dt.Rows.Add(r)
                MarkDirty()
            End SyncLock
        End Sub

        Public Shared Sub RequestRevision(docId As Integer, punchlistNotes As String, returnSection As String, staffName As String, Optional staffUserId As Integer = 1)
            Dim extCn As String = ""
            SyncLock _syncLock
                Dim doc = FindDocumentRow(docId)
                If doc Is Nothing Then Return
                RefuseTerminal(doc, "returned for revision")

                doc("CurrentStatus") = "FOR_REVISION"
                Dim prevPunchlist As String = doc("RevisionPunchlist").ToString()
                Dim stamp As String = "[" & DateTime.Now.ToString("yyyy-MM-dd HH:mm") & "] " & punchlistNotes
                doc("RevisionPunchlist") = If(String.IsNullOrWhiteSpace(prevPunchlist), stamp, prevPunchlist & vbCrLf & stamp)

                Dim targetSec As String = If(Not String.IsNullOrWhiteSpace(returnSection), returnSection, DocumentService.GetDefaultSectionForCategory(doc("DocType").ToString()))
                doc("AssignedSection") = targetSec
                doc("LastActionTaken") = "Revision requested by " & staffName & ": " & punchlistNotes
                If doc.Table.Columns.Contains("ModifiedByUserID") Then doc("ModifiedByUserID") = staffUserId
                MarkDirty()

                AddDirective(docId, "REVISION_REQUESTED", targetSec, punchlistNotes, staffName, isOffline:=True, issuedByUserId:=staffUserId)
                AddRoutingLog(docId, "Office of the Secretary-General", targetSec, staffName, "REVISION_REQUESTED", punchlistNotes, isOffline:=True, routedByUserId:=staffUserId)
                LogAudit(staffName, "Revision requested for Document #" & docId & " with punchlist: " & punchlistNotes, isOffline:=True, userId:=staffUserId)

                extCn = RowString(doc, "ExternalControlNumber")
            End SyncLock

            PushPortalStatusFor(extCn, "FOR_REVISION")
        End Sub

        Public Shared Sub ResubmitDocument(docId As Integer, staffName As String, notes As String, Optional staffUserId As Integer = 1)
            Dim extCn As String = ""
            SyncLock _syncLock
                Dim doc = FindDocumentRow(docId)
                If doc Is Nothing Then Return
                RefuseTerminal(doc, "resubmitted")

                Dim originSec As String = doc("AssignedSection").ToString()
                doc("CurrentStatus") = "FOR_REVIEW"
                doc("AssignedSection") = "Secretary-General"
                doc("LastActionTaken") = "Resubmitted for Sec Gen review: " & notes
                If doc.Table.Columns.Contains("ModifiedByUserID") Then doc("ModifiedByUserID") = staffUserId
                MarkDirty()

                AddRoutingLog(docId, originSec, "Office of the Secretary-General", staffName, "RESUBMITTED", notes, isOffline:=True, routedByUserId:=staffUserId)
                LogAudit(staffName, "Document #" & docId & " resubmitted by " & originSec & " to Sec Gen: " & notes, isOffline:=True, userId:=staffUserId)

                extCn = RowString(doc, "ExternalControlNumber")
            End SyncLock

            PushPortalStatusFor(extCn, "FOR_REVIEW")
        End Sub

        Public Shared Sub ApproveDocument(docId As Integer, staffName As String, notes As String, Optional staffUserId As Integer = 1)
            Dim extCn As String = ""
            SyncLock _syncLock
                Dim doc = FindDocumentRow(docId)
                If doc Is Nothing Then Return
                RefuseTerminal(doc, "approved")

                doc("CurrentStatus") = "APPROVED"
                doc("AssignedSection") = "Records Section"
                doc("LastActionTaken") = "Approved by " & staffName & ": " & notes
                If doc.Table.Columns.Contains("ModifiedByUserID") Then doc("ModifiedByUserID") = staffUserId
                MarkDirty()

                AddDirective(docId, "APPROVED", "Records Section", notes, staffName, isOffline:=True, issuedByUserId:=staffUserId)
                AddRoutingLog(docId, "Office of the Secretary-General", "Records Section", staffName, "APPROVED", notes, isOffline:=True, routedByUserId:=staffUserId)
                LogAudit(staffName, "Document #" & docId & " approved by " & staffName & ": " & notes, isOffline:=True, userId:=staffUserId)

                extCn = RowString(doc, "ExternalControlNumber")
            End SyncLock

            PushPortalStatusFor(extCn, "APPROVED")
        End Sub

        Public Shared Sub ReleaseDocument(docId As Integer, staffName As String, notes As String, Optional staffUserId As Integer = 1)
            Dim extCn As String = ""
            SyncLock _syncLock
                Dim doc = FindDocumentRow(docId)
                If doc Is Nothing Then Return
                RefuseTerminal(doc, "released")

                Dim dest As String = doc("DestinationOffice").ToString()
                doc("CurrentStatus") = "RELEASED"
                doc("AssignedSection") = "Archived / Released"
                doc("LastActionTaken") = "Released to " & dest & ": " & notes
                If doc.Table.Columns.Contains("ModifiedByUserID") Then doc("ModifiedByUserID") = staffUserId
                MarkDirty()

                AddRoutingLog(docId, "Records Section", dest, staffName, "RELEASED", notes, isOffline:=True, routedByUserId:=staffUserId)
                LogAudit(staffName, "Document #" & docId & " released: " & notes, isOffline:=True, userId:=staffUserId)

                extCn = RowString(doc, "ExternalControlNumber")
            End SyncLock

            PushPortalStatusFor(extCn, "RELEASED")
        End Sub

        Private Shared Function FindDocumentRow(docId As Integer) As DataRow
            EnsureInitialized()
            Dim dt = DataSet.Tables("Documents")
            If dt Is Nothing Then Return Nothing
            Dim rows = dt.Select("DocumentID = " & docId)
            If rows.Length = 0 Then Return Nothing
            Return rows(0)
        End Function

        ''' <summary>
        ''' The offline half of the terminal-state rule (the connected path refuses in
        ''' RoutingService/DirectiveService): a filed, released, archived, or completed
        ''' document cannot re-enter the workflow through a revision, resubmit, approval,
        ''' or release. Re-log of the same terminal status stays legal so the internal
        ''' transitions that re-record their own target do not refuse themselves.
        ''' </summary>
        Private Shared Sub RefuseTerminal(doc As DataRow, verb As String)
            Dim currentCode = doc("CurrentStatus").ToString().Trim().ToUpperInvariant()
            If DocumentStatus.IsTerminalStatus(currentCode) Then
                Throw New InvalidOperationException("This document is already " & DocumentStatus.DisplayName(currentCode) & " and can no longer be " & verb & ".")
            End If
        End Sub

        ' The portal push runs after _syncLock releases: it performs network I/O and must not
        ' hold the cache lock across it.
        Private Shared Sub PushPortalStatusFor(externalControlNumber As String, statusCode As String)
            If Not String.IsNullOrWhiteSpace(externalControlNumber) Then
                DocumentService.PushPortalStatusSafe(externalControlNumber, statusCode)
            End If
        End Sub

    End Class
End Namespace
