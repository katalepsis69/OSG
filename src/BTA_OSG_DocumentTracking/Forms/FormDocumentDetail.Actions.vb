Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Drawing
Imports System.Windows.Forms

Namespace BTA_OSG
    Partial Public Class FormDocumentDetail
        Private Sub UpdateActionButtons()
            Dim user = MainFrm.CurrentUser
            If user Is Nothing Then
                btnRequestRevision.Visible = False
                btnResubmit.Visible = False
                btnApprove.Visible = False
                btnRelease.Visible = False
                btnRoute.Enabled = False
                btnMove.Enabled = False
                Return
            End If

            Dim canRouteUser As Boolean = True
            If user.Table.Columns.Contains("CanRoute") AndAlso Not IsDBNull(user("CanRoute")) Then
                canRouteUser = Convert.ToBoolean(user("CanRoute"))
            End If

            Dim canMoveUser As Boolean = True
            If user.Table.Columns.Contains("CanMove") AndAlso Not IsDBNull(user("CanMove")) Then
                canMoveUser = Convert.ToBoolean(user("CanMove"))
            End If

            ' Soft-copy access is deny-by-default: only an explicit True grants it. The old
            ' allow-unless-False shape failed open exactly when the session was in the
            ' anomalous state (dropped mid-dialog, NULL flag) where enforcement matters.
            Dim canSoftCopyUser As Boolean = False
            If user.Table.Columns.Contains("CanSoftCopy") AndAlso Not IsDBNull(user("CanSoftCopy")) Then
                canSoftCopyUser = Convert.ToBoolean(user("CanSoftCopy"))
            End If

            Dim role = user("Role").ToString()
            Dim office = If(user.Table.Columns.Contains("Office") AndAlso Not IsDBNull(user("Office")), user("Office").ToString(), "")
            Dim isManager As Boolean = (role = "Secretary-General" OrElse role = "System Administrator" OrElse role = "OSG Chief")
            Dim status = DocRow("CurrentStatus").ToString().ToUpperInvariant()
            ' Routing is role-gated AND status-gated: a filed, released, archived, or
            ' completed document has left the active workflow and cannot be re-routed.
            btnRoute.Enabled = canRouteUser AndAlso Not DocumentStatus.IsTerminalStatus(status)
            btnMove.Enabled = canMoveUser
            btnLaunchPdf.Enabled = canSoftCopyUser

            Dim assignedSec = DocRow("AssignedSection").ToString()
            Dim userSection = If(Not String.IsNullOrWhiteSpace(office), office, role)

            ' Sec Gen / Manager can request revision while the document is still in the
            ' active workflow; terminal states refuse server-side too, so the button goes.
            btnRequestRevision.Visible = isManager AndAlso Not DocumentStatus.IsTerminalStatus(status)

            ' Sec Gen / Manager can approve if document is under review or received
            btnApprove.Visible = isManager AndAlso (status = "FOR_REVIEW" OrElse status.Contains("REVIEW") OrElse status = "RECEIVED")

            ' Resubmit is visible if status is FOR_REVISION and user belongs to assigned section or is manager
            Dim canResubmit As Boolean = (status = "FOR_REVISION" OrElse status.Contains("REVISION")) AndAlso (isManager OrElse role = assignedSec OrElse userSection = assignedSec)
            btnResubmit.Visible = canResubmit

            ' Release is visible if approved and user is Records Section or manager
            Dim canRelease As Boolean = (status = "APPROVED") AndAlso (isManager OrElse role = "Records Section" OrElse userSection = "Records Section")
            btnRelease.Visible = canRelease
        End Sub

        Private Function GetCurrentDocumentObject() As Document
            Return EmbeddedDB.MapRowToDocument(DocRow)
        End Function

        ''' <summary>
        ''' The held DocRow can be detached under the dialog by a sync tick that removes or
        ''' re-keys the row; touching it then throws RowNotInTableException out of the click
        ''' handler. Every action validates first and closes the dialog when the document
        ''' is gone, mirroring what OpenSelectedDocumentDetail does before opening.
        ''' </summary>
        Private Function EnsureDocRowLive() As Boolean
            If DocRow Is Nothing OrElse DocRow.RowState = System.Data.DataRowState.Deleted OrElse DocRow.RowState = System.Data.DataRowState.Detached OrElse EmbeddedDB.GetDocumentByID(DocID) Is Nothing Then
                MessageBox.Show("This document is no longer available. It may have been removed or changed by another workstation.", "Document Unavailable", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.Close()
                Return False
            End If
            ' The buttons were computed when the dialog opened; a sync tick that moved the
            ' document (another seat approved it, a directive landed) must re-gate them
            ' now, or a still-visible button would drive a second transition.
            UpdateActionButtons()
            Return True
        End Function

        ''' <summary>
        ''' The status id this dialog is showing, for the server-side from-status guard.
        ''' The cache stores status codes while SQL compares ids, so resolve here. When a
        ''' connected write cannot verify the shown status (unknown code), the caller must
        ''' refuse instead of writing unguarded: 0 would mean "no guard" server-side.
        ''' </summary>
        Private Function ExpectedFromStatusId() As Integer
            If AppStartup.ReferenceDataRepo Is Nothing Then Return 0
            Dim status = AppStartup.ReferenceDataRepo.GetStatusByCode(DocRow("CurrentStatus").ToString())
            Return If(status IsNot Nothing, status.StatusID, 0)
        End Function

        ''' <summary>
        ''' Fail-closed for the connected transitions: when the server is reachable and the
        ''' dialog's status code does not resolve to a seeded id, the from-status guard would
        ''' degrade to "no guard", so the action is refused with a refresh hint instead.
        ''' </summary>
        Private Function StatusUnverifiableForConnectedWrite() As Boolean
            If Program.Coordinator Is Nothing OrElse Not Program.IsDatabaseConnected Then Return False
            If AppStartup.ReferenceDataRepo Is Nothing Then Return False
            Return ExpectedFromStatusId() = 0
        End Function

        ''' <summary>
        ''' True when the server's live row no longer carries the status this dialog was
        ''' opened with: another seat approved, released, or routed it while it sat open.
        ''' </summary>
        Private Function DocumentStateMovedOnServer() As Boolean
            If Program.Coordinator Is Nothing OrElse Not Program.IsDatabaseConnected Then Return False
            If AppStartup.DocumentRepo Is Nothing Then Return False
            Dim expected = ExpectedFromStatusId()
            If expected = 0 Then Return False
            Dim liveDoc = AppStartup.DocumentRepo.GetById(DocID)
            If liveDoc Is Nothing Then Return False
            Return liveDoc.StatusID <> expected
        End Function

        Private Sub WarnStateMoved(actionName As String)
            Program.Coordinator.RefreshFromServer("Documents")
            RefreshGrids()
            MessageBox.Show("This document's state changed on the server while you were viewing it. The view has been refreshed; check the current status before " & actionName & ".", "Document State Changed", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub

        Private Function GetCurrentRoutingLogsList() As List(Of RoutingLog)
            Return EmbeddedDB.GetRoutingLogsForDocument(DocID)
        End Function

        Private Function GetCurrentDirectivesList() As List(Of ActionDirective)
            Return EmbeddedDB.GetDirectivesForDocument(DocID)
        End Function

        Private Sub OnPrintRoutingSlip(sender As Object, e As EventArgs)
            If Not EnsureDocRowLive() Then Return
            Dim docObj = GetCurrentDocumentObject()
            Dim logs = GetCurrentRoutingLogsList()
            Dim directives = GetCurrentDirectivesList()
            RoutingSlipPrintService.ShowPreview(docObj, logs, directives, Me)
        End Sub

        Private Sub OnRequestRevision(sender As Object, e As EventArgs)
            If MainFrm.CurrentUser Is Nothing Then
                MessageBox.Show("Authentication Required.", "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            If Not EnsureDocRowLive() Then Return

            Using dlg As New FormRevisionDialog(DocID, DocRow("DocCode").ToString(), DocRow("Title").ToString(), DocRow("AssignedSection").ToString())
                If dlg.ShowDialog(Me) = DialogResult.OK Then
                    If StatusUnverifiableForConnectedWrite() Then
                        WarnStateMoved("requesting a revision")
                        Return
                    End If
                    Dim staffName = MainFrm.CurrentUser("FullName").ToString()
                    Dim staffUserId As Integer = 1
                    If MainFrm.CurrentUser.Table.Columns.Contains("UserID") AndAlso Not IsDBNull(MainFrm.CurrentUser("UserID")) Then
                        staffUserId = Convert.ToInt32(MainFrm.CurrentUser("UserID"))
                    End If
                    Try
                        If Program.Coordinator IsNot Nothing Then
                            Program.Coordinator.RequestRevision(DocID, dlg.PunchlistNotes, dlg.TargetSection, staffName, staffUserId)
                        Else
                            EmbeddedDB.RequestRevision(DocID, dlg.PunchlistNotes, dlg.TargetSection, staffName)
                        End If
                    Catch ex As Exception
                        MessageBox.Show(ex.Message, "Revision Not Logged", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        RefreshGrids()
                        Return
                    End Try
                    MainFrm.ReportStatus("Revision order logged and returned to " & dlg.TargetSection & ".")
                    RefreshGrids()
                    MainFrm.RefreshActiveTabGrid()
                End If
            End Using
        End Sub

        Private Sub OnResubmit(sender As Object, e As EventArgs)
            If MainFrm.CurrentUser Is Nothing Then
                MessageBox.Show("Authentication Required.", "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            If Not EnsureDocRowLive() Then Return
            If DocumentStateMovedOnServer() Then
                WarnStateMoved("resubmitting")
                Return
            End If

            Using promptDlg As New FormInputPrompt("Resubmit Document for Sec Gen Review", "Enter resubmission remarks or notes detailing amendments made:", "Punchlist requirements completed.", True)
                If promptDlg.ShowDialog(Me) = DialogResult.OK Then
                    Dim notes = promptDlg.PromptValue
                    If Not String.IsNullOrWhiteSpace(notes) Then
                        If StatusUnverifiableForConnectedWrite() Then
                            WarnStateMoved("resubmitting")
                            Return
                        End If
                        Dim staffName = MainFrm.CurrentUser("FullName").ToString()
                        Dim staffUserId As Integer = 1
                        If MainFrm.CurrentUser.Table.Columns.Contains("UserID") AndAlso Not IsDBNull(MainFrm.CurrentUser("UserID")) Then
                            staffUserId = Convert.ToInt32(MainFrm.CurrentUser("UserID"))
                        End If
                        Try
                            If Program.Coordinator IsNot Nothing Then
                                Program.Coordinator.ResubmitDocument(DocID, staffName, notes, staffUserId, ExpectedFromStatusId())
                            Else
                                EmbeddedDB.ResubmitDocument(DocID, staffName, notes)
                            End If
                        Catch ex As Exception
                            MessageBox.Show(ex.Message, "Resubmit Not Logged", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                            RefreshGrids()
                            Return
                        End Try
                        MainFrm.ReportStatus("Document resubmitted to the Secretary-General for review.")
                        RefreshGrids()
                        MainFrm.RefreshActiveTabGrid()
                    End If
                End If
            End Using
        End Sub

        Private Sub OnApprove(sender As Object, e As EventArgs)
            If MainFrm.CurrentUser Is Nothing Then
                MessageBox.Show("Authentication Required.", "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            If Not EnsureDocRowLive() Then Return

            Dim confirm = MessageBox.Show("Approve document " & DocRow("DocCode").ToString() & " for official transmittal?", "Confirm Approval", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If confirm = DialogResult.Yes Then
                Dim staffName = MainFrm.CurrentUser("FullName").ToString()
                ' Attribute the action to the actual role: OSG Chief and System Administrator
                ' approvals must not be recorded as Secretary-General approvals.
                Dim approverRole As String = "OSG Staff"
                If MainFrm.CurrentUser.Table.Columns.Contains("Role") AndAlso Not IsDBNull(MainFrm.CurrentUser("Role")) Then
                    approverRole = MainFrm.CurrentUser("Role").ToString()
                End If
                Dim notes = "Approved by " & approverRole & "."
                Dim staffUserId As Integer = 1
                If MainFrm.CurrentUser.Table.Columns.Contains("UserID") AndAlso Not IsDBNull(MainFrm.CurrentUser("UserID")) Then
                    staffUserId = Convert.ToInt32(MainFrm.CurrentUser("UserID"))
                End If
                If DocumentStateMovedOnServer() Then
                    WarnStateMoved("approving")
                    Return
                End If
                If StatusUnverifiableForConnectedWrite() Then
                    WarnStateMoved("approving")
                    Return
                End If
                Try
                    If Program.Coordinator IsNot Nothing Then
                        Program.Coordinator.ApproveDocument(DocID, staffName, notes, staffUserId, ExpectedFromStatusId())
                    Else
                        EmbeddedDB.ApproveDocument(DocID, staffName, notes)
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message, "Approval Not Logged", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    RefreshGrids()
                    Return
                End Try
                MainFrm.ReportStatus("Document approved and routed to Records Section for release.")
                RefreshGrids()
                MainFrm.RefreshActiveTabGrid()
            End If
        End Sub

        Private Sub OnRelease(sender As Object, e As EventArgs)
            If MainFrm.CurrentUser Is Nothing Then
                MessageBox.Show("Authentication Required.", "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            If Not EnsureDocRowLive() Then Return

            Dim confirm = MessageBox.Show("Release document " & DocRow("DocCode").ToString() & " to destination office " & DocRow("DestinationOffice").ToString() & "?", "Confirm Release", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If confirm = DialogResult.Yes Then
                Dim staffName = MainFrm.CurrentUser("FullName").ToString()
                Dim staffUserId As Integer = 1
                If MainFrm.CurrentUser.Table.Columns.Contains("UserID") AndAlso Not IsDBNull(MainFrm.CurrentUser("UserID")) Then
                    staffUserId = Convert.ToInt32(MainFrm.CurrentUser("UserID"))
                End If
                If DocumentStateMovedOnServer() Then
                    WarnStateMoved("releasing")
                    Return
                End If
                If StatusUnverifiableForConnectedWrite() Then
                    WarnStateMoved("releasing")
                    Return
                End If
                Try
                    If Program.Coordinator IsNot Nothing Then
                        Program.Coordinator.ReleaseDocument(DocID, staffName, "Released to destination office.", staffUserId, ExpectedFromStatusId())
                    Else
                        EmbeddedDB.ReleaseDocument(DocID, staffName, "Released to destination office.")
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message, "Release Not Logged", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    RefreshGrids()
                    Return
                End Try
                MainFrm.ReportStatus("Document released and logged in the archive.")
                RefreshGrids()
                MainFrm.RefreshActiveTabGrid()
            End If
        End Sub

        Private Sub OnRouteDocument(sender As Object, e As EventArgs)
            If MainFrm.CurrentUser Is Nothing Then
                MessageBox.Show("Authentication Required.", "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            If MainFrm.CurrentUser.Table.Columns.Contains("CanRoute") AndAlso Not IsDBNull(MainFrm.CurrentUser("CanRoute")) AndAlso Not Convert.ToBoolean(MainFrm.CurrentUser("CanRoute")) Then
                MessageBox.Show("Access Denied: Current user does not have permission to route documents.", "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            If Not EnsureDocRowLive() Then Return

            Dim staffUserId As Integer = 1
            If MainFrm.CurrentUser IsNot Nothing AndAlso MainFrm.CurrentUser.Table.Columns.Contains("UserID") AndAlso Not IsDBNull(MainFrm.CurrentUser("UserID")) Then
                staffUserId = Convert.ToInt32(MainFrm.CurrentUser("UserID"))
            End If

            Using dlg As New FormRouteDocument(DocID, DocRow("DestinationOffice").ToString(), MainFrm.CurrentUser("FullName").ToString(), staffUserId)
                If dlg.ShowDialog(Me) = DialogResult.OK Then
                    ' Reload the row first: the coordinator path may have transitioned the
                    ' status server-side and pulled it back, and the portal must hear the
                    ' document's real status, not an assumption about what routing did.
                    LoadDocData()
                    If DocRow.Table.Columns.Contains("ExternalControlNumber") AndAlso Not IsDBNull(DocRow("ExternalControlNumber")) Then
                        Dim extCn = DocRow("ExternalControlNumber").ToString()
                        If Not String.IsNullOrWhiteSpace(extCn) Then
                            DocumentService.PushPortalStatusSafe(extCn, DocRow("CurrentStatus").ToString())
                        End If
                    End If
                    RefreshGrids()
                    MainFrm.RefreshActiveTabGrid()
                End If
            End Using
        End Sub

        Private Sub OnMoveStorage(sender As Object, e As EventArgs)
            If MainFrm.CurrentUser Is Nothing Then
                MessageBox.Show("Authentication Required.", "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            If Not EnsureDocRowLive() Then Return

            If MainFrm.CurrentUser.Table.Columns.Contains("CanMove") AndAlso Not IsDBNull(MainFrm.CurrentUser("CanMove")) AndAlso Not Convert.ToBoolean(MainFrm.CurrentUser("CanMove")) Then
                MessageBox.Show("Access Denied: Current user does not have permission to transfer document storage locations.", "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim staffUserId As Integer = 1
            If MainFrm.CurrentUser IsNot Nothing AndAlso MainFrm.CurrentUser.Table.Columns.Contains("UserID") AndAlso Not IsDBNull(MainFrm.CurrentUser("UserID")) Then
                staffUserId = Convert.ToInt32(MainFrm.CurrentUser("UserID"))
            End If

            Dim currentLoc = String.Format("{0}/{1}/{2}", DocRow("CabinetID"), DocRow("ShelfNo"), DocRow("BoxCode"))
            Using dlg As New FormMoveStorage(DocID, currentLoc, MainFrm.CurrentUser("FullName").ToString(), staffUserId)
                If dlg.ShowDialog(Me) = DialogResult.OK Then
                    RefreshGrids()
                    MainFrm.RefreshActiveTabGrid()
                End If
            End Using
        End Sub

        Private Sub OnLaunchPDF(sender As Object, e As EventArgs)
            ' Deny-by-default: no session, missing column, or NULL flag all refuse.
            If MainFrm.CurrentUser Is Nothing Then
                MessageBox.Show("Authentication Required.", "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            Dim canSoftCopy As Boolean = False
            If MainFrm.CurrentUser.Table.Columns.Contains("CanSoftCopy") AndAlso Not IsDBNull(MainFrm.CurrentUser("CanSoftCopy")) Then
                canSoftCopy = Convert.ToBoolean(MainFrm.CurrentUser("CanSoftCopy"))
            End If
            If Not canSoftCopy Then
                MessageBox.Show("Access Denied: Current user does not have permission to view or launch soft-copy attachments.", "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            If Not EnsureDocRowLive() Then Return

            Dim url = DocRow("GDriveURL").ToString()
            If String.IsNullOrWhiteSpace(url) Then
                MessageBox.Show("No soft-copy attachment is linked to this document.", "No Attachment", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If
            Dim errUrl As String = ""
            If Not EmbeddedDB.ValidateGDriveURL(url, errUrl) Then
                MessageBox.Show(errUrl, "Security Validation Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            Try
                Process.Start(New ProcessStartInfo With {.FileName = url, .UseShellExecute = True})
                If MainFrm.CurrentUser IsNot Nothing Then
                    EmbeddedDB.LogAudit(MainFrm.CurrentUser("FullName").ToString(), "Launched Document Soft Copy: " & url, actionType:="SOFT_COPY_LAUNCHED")
                End If
            Catch ex As Exception
                MessageBox.Show("Error opening document: " & ex.Message, "Launch Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub
    End Class
End Namespace
