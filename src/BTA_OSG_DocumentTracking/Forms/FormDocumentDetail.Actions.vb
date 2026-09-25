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

            Dim canSoftCopyUser As Boolean = True
            If user.Table.Columns.Contains("CanSoftCopy") AndAlso Not IsDBNull(user("CanSoftCopy")) Then
                canSoftCopyUser = Convert.ToBoolean(user("CanSoftCopy"))
            End If

            btnRoute.Enabled = canRouteUser
            btnMove.Enabled = canMoveUser
            btnLaunchPdf.Enabled = canSoftCopyUser

            Dim role = user("Role").ToString()
            Dim office = If(user.Table.Columns.Contains("Office") AndAlso Not IsDBNull(user("Office")), user("Office").ToString(), "")
            Dim isManager As Boolean = (role = "Secretary-General" OrElse role = "System Administrator" OrElse role = "OSG Chief")
            Dim status = DocRow("CurrentStatus").ToString().ToUpperInvariant()
            Dim assignedSec = DocRow("AssignedSection").ToString()
            Dim userSection = If(Not String.IsNullOrWhiteSpace(office), office, role)

            ' Sec Gen / Manager can request revision if document is not released/filed
            btnRequestRevision.Visible = isManager AndAlso (status <> "RELEASED" AndAlso status <> "FILED")

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

        Private Function GetCurrentRoutingLogsList() As List(Of RoutingLog)
            Return EmbeddedDB.GetRoutingLogsForDocument(DocID)
        End Function

        Private Function GetCurrentDirectivesList() As List(Of ActionDirective)
            Return EmbeddedDB.GetDirectivesForDocument(DocID)
        End Function

        Private Sub OnPrintRoutingSlip(sender As Object, e As EventArgs)
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

            Using dlg As New FormRevisionDialog(DocID, DocRow("DocCode").ToString(), DocRow("Title").ToString(), DocRow("AssignedSection").ToString())
                If dlg.ShowDialog(Me) = DialogResult.OK Then
                    Dim staffName = MainFrm.CurrentUser("FullName").ToString()
                    Dim staffUserId As Integer = 1
                    If MainFrm.CurrentUser.Table.Columns.Contains("UserID") AndAlso Not IsDBNull(MainFrm.CurrentUser("UserID")) Then
                        staffUserId = Convert.ToInt32(MainFrm.CurrentUser("UserID"))
                    End If
                    If Program.Coordinator IsNot Nothing Then
                        Program.Coordinator.RequestRevision(DocID, dlg.PunchlistNotes, dlg.TargetSection, staffName, staffUserId)
                    Else
                        EmbeddedDB.RequestRevision(DocID, dlg.PunchlistNotes, dlg.TargetSection, staffName)
                    End If
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

            Using promptDlg As New FormInputPrompt("Resubmit Document for Sec Gen Review", "Enter resubmission remarks or notes detailing amendments made:", "Punchlist requirements completed.", True)
                If promptDlg.ShowDialog(Me) = DialogResult.OK Then
                    Dim notes = promptDlg.PromptValue
                    If Not String.IsNullOrWhiteSpace(notes) Then
                        Dim staffName = MainFrm.CurrentUser("FullName").ToString()
                        Dim staffUserId As Integer = 1
                        If MainFrm.CurrentUser.Table.Columns.Contains("UserID") AndAlso Not IsDBNull(MainFrm.CurrentUser("UserID")) Then
                            staffUserId = Convert.ToInt32(MainFrm.CurrentUser("UserID"))
                        End If
                        If Program.Coordinator IsNot Nothing Then
                            Program.Coordinator.ResubmitDocument(DocID, staffName, notes, staffUserId)
                        Else
                            EmbeddedDB.ResubmitDocument(DocID, staffName, notes)
                        End If
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

            Dim confirm = MessageBox.Show("Approve document " & DocRow("DocCode").ToString() & " for official transmittal?", "Confirm Approval", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If confirm = DialogResult.Yes Then
                Dim staffName = MainFrm.CurrentUser("FullName").ToString()
                Dim staffUserId As Integer = 1
                If MainFrm.CurrentUser.Table.Columns.Contains("UserID") AndAlso Not IsDBNull(MainFrm.CurrentUser("UserID")) Then
                    staffUserId = Convert.ToInt32(MainFrm.CurrentUser("UserID"))
                End If
                If Program.Coordinator IsNot Nothing Then
                    Program.Coordinator.ApproveDocument(DocID, staffName, "Approved by Secretary-General.", staffUserId)
                Else
                    EmbeddedDB.ApproveDocument(DocID, staffName, "Approved by Secretary-General.")
                End If
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

            Dim confirm = MessageBox.Show("Release document " & DocRow("DocCode").ToString() & " to destination office " & DocRow("DestinationOffice").ToString() & "?", "Confirm Release", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If confirm = DialogResult.Yes Then
                Dim staffName = MainFrm.CurrentUser("FullName").ToString()
                Dim staffUserId As Integer = 1
                If MainFrm.CurrentUser.Table.Columns.Contains("UserID") AndAlso Not IsDBNull(MainFrm.CurrentUser("UserID")) Then
                    staffUserId = Convert.ToInt32(MainFrm.CurrentUser("UserID"))
                End If
                If Program.Coordinator IsNot Nothing Then
                    Program.Coordinator.ReleaseDocument(DocID, staffName, "Released to destination office.", staffUserId)
                Else
                    EmbeddedDB.ReleaseDocument(DocID, staffName, "Released to destination office.")
                End If
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

            Dim staffUserId As Integer = 1
            If MainFrm.CurrentUser IsNot Nothing AndAlso MainFrm.CurrentUser.Table.Columns.Contains("UserID") AndAlso Not IsDBNull(MainFrm.CurrentUser("UserID")) Then
                staffUserId = Convert.ToInt32(MainFrm.CurrentUser("UserID"))
            End If

            Using dlg As New FormRouteDocument(DocID, DocRow("DestinationOffice").ToString(), MainFrm.CurrentUser("FullName").ToString(), staffUserId)
                If dlg.ShowDialog(Me) = DialogResult.OK Then
                    If DocRow.Table.Columns.Contains("ExternalControlNumber") AndAlso Not IsDBNull(DocRow("ExternalControlNumber")) Then
                        Dim extCn = DocRow("ExternalControlNumber").ToString()
                        If Not String.IsNullOrWhiteSpace(extCn) Then
                            DocumentService.PushPortalStatusSafe(extCn, "FOR_REVIEW")
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
            If MainFrm.CurrentUser IsNot Nothing AndAlso MainFrm.CurrentUser.Table.Columns.Contains("CanSoftCopy") AndAlso Not IsDBNull(MainFrm.CurrentUser("CanSoftCopy")) AndAlso Not Convert.ToBoolean(MainFrm.CurrentUser("CanSoftCopy")) Then
                MessageBox.Show("Access Denied: Current user does not have permission to view or launch soft-copy attachments.", "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim url = DocRow("GDriveURL").ToString()
            Dim errUrl As String = ""
            If Not EmbeddedDB.ValidateGDriveURL(url, errUrl) Then
                MessageBox.Show(errUrl, "Security Validation Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            Try
                Process.Start(New ProcessStartInfo With {.FileName = url, .UseShellExecute = True})
                If MainFrm.CurrentUser IsNot Nothing Then
                    EmbeddedDB.LogAudit(MainFrm.CurrentUser("FullName").ToString(), "Launched Document Soft Copy: " & url)
                End If
            Catch ex As Exception
                MessageBox.Show("Error opening document: " & ex.Message, "Launch Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub
    End Class
End Namespace
