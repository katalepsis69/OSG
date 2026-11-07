Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Windows.Forms

Namespace BTA_OSG
    ' Shared document-grid behavior: keyboard open (Enter = full details, Shift+Enter =
    ' quick view), the bulk routing entry point, and the persisted dashboard filters.
    Partial Public Class FormMain
        ''' <summary>
        ''' Document grids open on the keyboard the way the Portal queue already does:
        ''' Enter opens the full detail dialog, Shift+Enter the quick view. Suppressing
        ''' the keypress stops the grid from swallowing Enter as a row move.
        ''' </summary>
        Private Sub HandleDocumentGridKeys(dgv As DataGridView, e As KeyEventArgs)
            If e.KeyCode <> Keys.Enter Then Return
            e.Handled = True
            e.SuppressKeyPress = True
            If e.Shift Then
                OpenSelectedDocumentQuickView(dgv)
            Else
                OpenSelectedDocumentDetail(dgv)
            End If
        End Sub

        Private Sub OpenSelectedDocumentQuickView(dgv As DataGridView)
            If dgv Is Nothing OrElse dgv.CurrentRow Is Nothing OrElse dgv.CurrentRow.Cells("DocumentID").Value Is Nothing OrElse IsDBNull(dgv.CurrentRow.Cells("DocumentID").Value) Then
                ReportStatus("Please select a document row from the grid first.")
                Return
            End If
            Dim docId = CInt(dgv.CurrentRow.Cells("DocumentID").Value)
            ' Same guard as OpenSelectedDocumentDetail: the row can vanish between grid
            ' paint and keypress, and the quick view constructor would dispose itself,
            ' turning ShowDialog into an ObjectDisposedException on a dead object.
            If EmbeddedDB.GetDocumentByID(docId) Is Nothing Then
                ReportStatus("This document is no longer available. Refresh the grid and try again.")
                Return
            End If
            Using dlg As New FormDocumentQuickView(docId)
                ' Full Details hands back OK so the complete dialog opens from the same
                ' seat the quick view was launched on.
                If dlg.ShowDialog(Me) = DialogResult.OK Then
                    OpenSelectedDocumentDetail(dgv)
                End If
            End Using
        End Sub

        ''' <summary>
        ''' Route every selected document through one dialog: destination, action, and
        ''' remarks are asked once; each document's From office is its current holding
        ''' section, so custody never records a wrong leg. One dialog, per-document
        ''' guards and reporting.
        ''' </summary>
        Private Sub RouteSelectedDocuments(dgv As DataGridView)
            If CurrentUser Is Nothing Then
                ReportRefusal("Authentication Required: Tap RFID card to route documents.")
                Return
            End If
            If CurrentUser.Table.Columns.Contains("CanRoute") AndAlso Not IsDBNull(CurrentUser("CanRoute")) AndAlso Not Convert.ToBoolean(CurrentUser("CanRoute")) Then
                ReportRefusal("Access Denied: Current user does not have permission to route documents.")
                Return
            End If

            Dim docIds As New List(Of Integer)()
            For Each row As DataGridViewRow In dgv.SelectedRows
                If row.Cells("DocumentID").Value IsNot Nothing AndAlso Not IsDBNull(row.Cells("DocumentID").Value) Then
                    docIds.Add(CInt(row.Cells("DocumentID").Value))
                End If
            Next
            If docIds.Count = 0 Then
                ReportStatus("Select one or more document rows to route.")
                Return
            End If

            ' A vanished row (sync delete, refilter) drops out before the dialog opens.
            docIds.RemoveAll(Function(id) EmbeddedDB.GetDocumentByID(id) Is Nothing)
            If docIds.Count = 0 Then
                ReportStatus("The selected documents are no longer available. Refresh the grid and try again.")
                Return
            End If

            Dim staffUserId As Integer = CurrentStaffUserId()

            Using dlg As New FormRouteDocument(docIds(0), "", CurrentUser("FullName").ToString(), staffUserId, OfficeHistoryForRouting(), LastRoutedOfficeFor(docIds(0)), docIds)
                If dlg.ShowDialog(Me) = DialogResult.OK Then
                    RefreshActiveTabGrid()
                End If
            End Using
        End Sub

        Friend Function OfficeHistoryForRouting() As List(Of String)
            ' Destination and origin history cover the two route-dialog fields clerks
            ' actually retype; one combined list keeps the autocomplete simple.
            Dim offices As New List(Of String)()
            Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each v As String In DistinctDocumentColumn("DestinationOffice")
                If seen.Add(v) Then offices.Add(v)
            Next
            For Each v As String In DistinctDocumentColumn("OriginatingOffice")
                If seen.Add(v) Then offices.Add(v)
            Next
            Return offices
        End Function

        ''' <summary>
        ''' The most recent To office this document was routed to, or "" when it has never
        ''' moved: prefilling the route dialog with the last hop is the lazy 20% of saved
        ''' routing templates.
        ''' </summary>
        Friend Function LastRoutedOfficeFor(docId As Integer) As String
            Dim logs = EmbeddedDB.GetRoutingLogsForDocument(docId)
            Return If(logs.Count > 0, logs(logs.Count - 1).ToOffice, "")
        End Function

        Private Sub SaveDashboardUiState()
            Try
                Dim ui = AppSettings.Instance.UiStateSettings
                If ui Is Nothing Then
                    ui = New UiStateSettings()
                    AppSettings.Instance.UiStateSettings = ui
                End If
                ui.DashboardSection = If(cmbDashSection.SelectedItem, "").ToString()
                ui.DashboardCategory = If(cmbDashCategory.SelectedItem, "").ToString()
                ui.DashboardStatus = If(cmbDashStatus.SelectedItem, "").ToString()
                ui.DashboardOnlyMine = chkDashOnlyMine.Checked
                AppSettings.Save(AppSettings.Instance)
            Catch ex As Exception
                System.Diagnostics.Trace.TraceWarning("Dashboard filters could not be saved: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Restores last session's dashboard filters, but only entries the live lists
        ''' still carry: a renamed section desk fails open to "All" instead of filtering
        ''' the grid to nothing.
        ''' </summary>
        Private Sub ApplySavedDashboardUiState()
            Dim ui = AppSettings.Instance.UiStateSettings
            If ui Is Nothing Then Return
            If cmbDashSection Is Nothing Then Return
            If ui.DashboardSection.Length > 0 AndAlso cmbDashSection.Items.Contains(ui.DashboardSection) Then
                cmbDashSection.SelectedItem = ui.DashboardSection
            End If
            If ui.DashboardCategory.Length > 0 AndAlso cmbDashCategory.Items.Contains(ui.DashboardCategory) Then
                cmbDashCategory.SelectedItem = ui.DashboardCategory
            End If
            If ui.DashboardStatus.Length > 0 AndAlso cmbDashStatus.Items.Contains(ui.DashboardStatus) Then
                cmbDashStatus.SelectedItem = ui.DashboardStatus
            End If
            chkDashOnlyMine.Checked = ui.DashboardOnlyMine
        End Sub
    End Class
End Namespace
