Option Explicit On
Option Strict On

Imports System
Imports System.Drawing
Imports System.Threading.Tasks
Imports System.Windows.Forms

Namespace BTA_OSG
    Partial Public Class FormMain
        Private Sub SetupPortalIntakeView()
            viewPortalIntake = New Panel With {.Dock = DockStyle.Fill, .BackColor = CivicCalmTheme.ColorCanvas}

            Dim tblPortal As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .RowCount = 2,
                .BackColor = CivicCalmTheme.ColorCanvas,
                .Padding = New Padding(0)
            }
            tblPortal.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            tblPortal.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            tblPortal.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

            Dim pnlPortalActionCard As New Panel With {
                .Dock = DockStyle.Fill,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .Margin = New Padding(0, 0, 0, 12),
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(12)
            }
            ApplyCardBorder(pnlPortalActionCard)

            Dim flwPortalActions As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .FlowDirection = FlowDirection.LeftToRight,
                .WrapContents = True,
                .BackColor = CivicCalmTheme.ColorSurface
            }

            btnPortalRefresh = New Button With {
                .Text = "&Refresh Queue",
                .Size = New Size(130, 34),
                .BackColor = CivicCalmTheme.ColorPrimary,
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(0, 6, 8, 0)
            }
            btnPortalRefresh.FlatAppearance.BorderSize = 0
            AddHandler btnPortalRefresh.Click, Sub() RefreshPortalQueue()

            btnPortalImportSelected = New Button With {
                .Text = "&Import Selected Submission",
                .Size = New Size(210, 34),
                .BackColor = CivicCalmTheme.ColorWell,
                .ForeColor = CivicCalmTheme.ColorInk,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(0, 6, 16, 0)
            }
            btnPortalImportSelected.FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder
            AddHandler btnPortalImportSelected.Click, AddressOf OnImportPortalSubmission

            lblPortalStatus = New Label With {
                .Text = "Portal Bridge: Standby",
                .UseMnemonic = False,
                .Font = CivicCalmTheme.FontBody,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .AutoSize = True,
                .Margin = New Padding(0, 12, 0, 0)
            }

            flwPortalActions.Controls.AddRange(New Control() {btnPortalRefresh, btnPortalImportSelected, lblPortalStatus})
            pnlPortalActionCard.Controls.Add(flwPortalActions)

            Dim pnlGridCard As New Panel With {
                .Dock = DockStyle.Fill,
                .Margin = New Padding(0),
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(12)
            }
            ApplyCardBorder(pnlGridCard)

            dgvPortalQueue = New DataGridView With {
                .Dock = DockStyle.Fill,
                .ReadOnly = True,
                .AllowUserToAddRows = False
            }
            DataGridStyler.ApplyCivicStyle(dgvPortalQueue)
            AddHandler dgvPortalQueue.CellDoubleClick, Sub(s, e)
                                                           If e.RowIndex >= 0 Then OnImportPortalSubmission(s, EventArgs.Empty)
                                                       End Sub
            AddHandler dgvPortalQueue.KeyDown, Sub(s, e)
                                                   If e.KeyCode = Keys.Enter Then
                                                       e.Handled = True
                                                       OnImportPortalSubmission(s, EventArgs.Empty)
                                                   End If
                                               End Sub

            lblPortalWatermark = CreateGridWatermark(pnlGridCard)

            pnlGridCard.Controls.Add(dgvPortalQueue)

            tblPortal.Controls.Add(pnlPortalActionCard, 0, 0)
            tblPortal.Controls.Add(pnlGridCard, 0, 1)

            viewPortalIntake.Controls.Add(tblPortal)
        End Sub

        Public Sub UpdatePortalIntakeBadgeAndQueue(items As List(Of PortalSubmission))
            If Me.IsDisposed Then Return
            _portalSubmissions = If(items, New List(Of PortalSubmission)())

            ' Update sidebar nav button for Portal Intake (Index 6)
            If navButtons IsNot Nothing AndAlso navButtons.Count > 6 Then
                Dim btnPortal = navButtons(6)
                Dim count = _portalSubmissions.Count
                If count > 0 Then
                    If Not sidebarCollapsed Then
                        btnPortal.Text = $"  Portal Intake ({count})"
                    End If
                    tipNav.SetToolTip(btnPortal, $"Portal Intake ({count} pending submission(s))")
                Else
                    If Not sidebarCollapsed Then
                        btnPortal.Text = "  Portal Intake"
                    End If
                    tipNav.SetToolTip(btnPortal, If(sidebarCollapsed, "Portal Intake", Nothing))
                End If
            End If

            ' If Portal Intake view is currently open, update its grid and status
            If activeNavIndex = 6 AndAlso dgvPortalQueue IsNot Nothing Then
                activeViewRetry = Nothing
                dgvPortalQueue.DataSource = Nothing
                dgvPortalQueue.DataSource = _portalSubmissions
                DataGridStyler.FormatPortalQueueColumns(dgvPortalQueue)

                If _portalSubmissions.Count = 0 Then
                    lblPortalStatus.Text = "Portal Bridge: Connected (0 pending)"
                    DataGridStyler.SetEmptyState(dgvPortalQueue, lblPortalWatermark, "No unimported public submissions pending in external queue.")
                Else
                    lblPortalStatus.Text = $"Portal Bridge: Connected ({_portalSubmissions.Count} pending)"
                    DataGridStyler.SetPopulatedState(dgvPortalQueue, lblPortalWatermark)
                End If

                lblStatusCount.Text = $"{_portalSubmissions.Count} Pending"
            End If
        End Sub

        Public Sub RefreshPortalQueue()
            If AppStartup.PortalBridgeClient Is Nothing OrElse Not AppSettings.Instance.PortalSettings.PortalEnabled Then
                lblPortalStatus.Text = "Portal Bridge: Disabled"
                DataGridStyler.SetEmptyState(dgvPortalQueue, lblPortalWatermark, "Public Portal Bridge is disabled in settings.")
                Return
            End If

            lblPortalStatus.Text = "Portal Bridge: Connecting"
            DataGridStyler.SetLoadingState(dgvPortalQueue, lblPortalWatermark)

            Task.Run(Async Function()
                         Try
                             Dim items = Await AppStartup.PortalBridgeClient.FetchExternalQueueAsync().ConfigureAwait(False)
                             If Me.IsDisposed Then Return

                             Me.Invoke(Sub()
                                           activeViewRetry = Nothing
                                           UpdatePortalIntakeBadgeAndQueue(items)
                                           lblStatusMessage.Text = $"Portal queue refreshed: {items.Count} pending external submission(s)."
                                       End Sub)
                         Catch ex As Exception
                             If Me.IsDisposed Then Return
                             Me.Invoke(Sub()
                                           lblPortalStatus.Text = "Portal Bridge: Offline / Unreachable"
                                           DataGridStyler.SetErrorState(dgvPortalQueue, lblPortalWatermark, "Unable to retrieve the external portal queue. (Alt+R or click this banner to retry)")
                                           activeViewRetry = AddressOf RefreshPortalQueue
                                           lblStatusMessage.Text = "External portal bridge query failed."
                                       End Sub)
                         End Try
                     End Function)
        End Sub

        Private Async Sub OnImportPortalSubmission(sender As Object, e As EventArgs)
            Dim submission As PortalSubmission = Nothing
            If Not ValidateSelectedPortalSubmission(submission) Then Return
            If Not ConfirmSubmissionImport(submission) Then Return
            Await ExecuteSubmissionImportAsync(submission)
        End Sub

        Private Function ValidateSelectedPortalSubmission(ByRef submission As PortalSubmission) As Boolean
            submission = Nothing
            If dgvPortalQueue.CurrentRow Is Nothing OrElse dgvPortalQueue.CurrentRow.DataBoundItem Is Nothing Then
                MessageBox.Show("Please select an external submission from the queue to import.", "Select Submission", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return False
            End If

            ' Role security check: RECORDS, SG, SYSADMIN
            If CurrentUser Is Nothing Then
                MessageBox.Show("RFID Badge Authentication is required before importing documents into the official registry.", "Authentication Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return False
            End If

            Dim role = CurrentUser("Role").ToString().ToUpperInvariant()
            Dim office = If(CurrentUser.Table.Columns.Contains("Office") AndAlso Not IsDBNull(CurrentUser("Office")), CurrentUser("Office").ToString().ToUpperInvariant(), "")
            Dim isAuthorized = role.Contains("RECORDS") OrElse office.Contains("RECORDS") OrElse role.Contains("SECRETARY-GENERAL") OrElse role.Contains("ADMIN") OrElse role.Contains("CHIEF")
            If Not isAuthorized Then
                MessageBox.Show("Access Denied: Only Records Section, Secretary-General, or System Administrator staff may import external submissions.", "Unauthorized", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                Return False
            End If

            submission = TryCast(dgvPortalQueue.CurrentRow.DataBoundItem, PortalSubmission)
            Return submission IsNot Nothing
        End Function

        Private Function ConfirmSubmissionImport(submission As PortalSubmission) As Boolean
            Dim confirm = MessageBox.Show(
                $"Import external submission into OSG Registry?" & vbCrLf & vbCrLf &
                $"Control Number: {submission.ControlNumber}" & vbCrLf &
                $"Title: {submission.DocumentTitle}" & vbCrLf &
                $"Submitter: {submission.RequesterName}" & vbCrLf &
                $"Category: {submission.Category}",
                "Confirm External Import",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question)

            Return confirm = DialogResult.Yes
        End Function

        ' Retrying a dead SQL host costs a connect timeout on every repository call inside
        ' RegisterImportedExternalDocument, so the app-level probe result acts as the circuit
        ' breaker: once SQL is known offline, imports register locally without paying stacked
        ' connect timeouts, and the one retry per app run executes off the UI thread.
        Private Async Function ExecuteSubmissionImportAsync(submission As PortalSubmission) As Task
            Dim alreadyRegistered = EmbeddedDB.DataSet.Tables("Documents").Select("ExternalControlNumber = '" & submission.ControlNumber.Replace("'", "''") & "'")
            If alreadyRegistered.Length > 0 Then
                MessageBox.Show($"Submission {submission.ControlNumber} is already in the registry as {alreadyRegistered(0)("DocCode")}.", "Already Imported", MessageBoxButtons.OK, MessageBoxIcon.Information)
                RefreshPortalQueue()
                Return
            End If

            Dim userId = CInt(CurrentUser("UserID"))
            Dim docCode = EmbeddedDB.GenerateDocCode(submission.Category)
            Dim origin = submission.RequesterName & " (" & submission.RequesterEmail & ")"
            Dim dest = "Office of the Secretary-General"
            Dim assignedSection = DocumentService.GetDefaultSectionForCategory(submission.Category)
            Dim lastAction = "Imported from Public Portal and routed to " & assignedSection

            If Program.IsDatabaseConnected Then
                Try
                    Dim importedDoc = Await Task.Run(Function() AppStartup.DocService.RegisterImportedExternalDocument(submission, userId))
                    docCode = importedDoc.DocCode
                    origin = importedDoc.OriginOffice
                    dest = importedDoc.DestinationOffice
                    assignedSection = importedDoc.AssignedSection
                    lastAction = importedDoc.LastActionTaken
                Catch ex As Microsoft.Data.SqlClient.SqlException
                    ' Fall back to local embedded database cache when SQL Server is unreachable.
                    Program.IsDatabaseConnected = False
                    lblStatusMessage.Text = "SQL Server unreachable, registering locally: " & ex.Message
                Catch ex As Exception
                    ' Other failures (mapping, logic) must not disable the SQL path for later imports.
                    lblStatusMessage.Text = "SQL registration failed, registering locally: " & ex.GetType().Name
                End Try
            End If

            EmbeddedDB.AddDocument(
                code:=docCode,
                docType:=submission.Category,
                title:=submission.DocumentTitle,
                origin:=origin,
                dest:=dest,
                cab:="CAB-A", shelf:="S-1", box:="BOX-01",
                url:="",
                status:="RECEIVED",
                assigned:=CurrentUser("FullName").ToString(),
                flowDirection:="INCOMING",
                assignedSection:=assignedSection,
                lastAction:=lastAction,
                externalControlNumber:=submission.ControlNumber)

            ' Acknowledged here rather than trusted from DocumentService's fire-and-forget so the
            ' result can be reported truthfully. The portal's ack is idempotent, so the extra call
            ' the SQL path makes is harmless.
            Dim portalNotified As Boolean = False
            If AppStartup.PortalBridgeClient IsNot Nothing Then
                Try
                    portalNotified = Await AppStartup.PortalBridgeClient.AcknowledgeImportAsync(submission.ControlNumber)
                Catch ackEx As Exception
                    lblStatusMessage.Text = "External portal acknowledgement failed: " & ackEx.Message
                End Try
            End If

            If portalNotified Then
                lblStatusMessage.Text = $"Imported {submission.ControlNumber} as {docCode}, routed to {assignedSection}."
            Else
                MessageBox.Show(
                    $"Document imported into the local registry, but the external portal did NOT acknowledge." & vbCrLf & vbCrLf &
                    $"Internal Doc Code: {docCode}" & vbCrLf &
                    $"External Control No: {submission.ControlNumber}" & vbCrLf & vbCrLf &
                    "Refresh the portal queue and import again if it still lists this submission.",
                    "Imported, Portal Not Notified",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning)
            End If

            ' RefreshPortalQueue already fetches the external queue; a second fetch via
            ' OnPortalPollTick would duplicate the HTTP round trip.
            RefreshActiveTabGrid()
            RefreshPortalQueue()
        End Function
    End Class
End Namespace
