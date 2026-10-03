Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Drawing
Imports System.Windows.Forms

Namespace BTA_OSG
    Partial Public Class FormMain
        Private Sub SetupRegistryView()
            viewRegistry = New Panel With {.Dock = DockStyle.Fill, .BackColor = CivicCalmTheme.ColorCanvas}

            Dim tblMain As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 2,
                .RowCount = 1
            }
            tblMain.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, Dpi(400.0F)))
            tblMain.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))

            ' Left: Form Card Panel
            Dim pnlFormCard As New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(16),
                .Margin = New Padding(0, 0, 16, 0)
            }
            ApplyCardBorder(pnlFormCard)

            Dim pnlFormScroll As New Panel With {
                .Dock = DockStyle.Fill,
                .AutoScroll = True,
                .BackColor = CivicCalmTheme.ColorSurface
            }

            txtTitle = New TextBox With {.Height = 28, .Text = "", .TabIndex = 1, .BackColor = CivicCalmTheme.ColorSurface, .ForeColor = CivicCalmTheme.ColorInk, .BorderStyle = BorderStyle.FixedSingle}

            cmbDocType = New ComboBox With {.DropDownStyle = ComboBoxStyle.DropDownList, .Height = 28, .TabIndex = 2, .BackColor = CivicCalmTheme.ColorSurface, .ForeColor = CivicCalmTheme.ColorInk}
            cmbDocType.Items.AddRange(New Object() {"Regular Communication", "Legislative", "Finance", "Travel Order"})
            cmbDocType.SelectedIndex = 0

            cmbFlowDirection = New ComboBox With {.DropDownStyle = ComboBoxStyle.DropDownList, .Height = 28, .TabIndex = 3, .BackColor = CivicCalmTheme.ColorSurface, .ForeColor = CivicCalmTheme.ColorInk}
            cmbFlowDirection.Items.AddRange(New Object() {"Incoming", "Outgoing"})
            cmbFlowDirection.SelectedIndex = 0

            cmbDeadlinePreset = New ComboBox With {.DropDownStyle = ComboBoxStyle.DropDownList, .Height = 28, .TabIndex = 4, .BackColor = CivicCalmTheme.ColorSurface, .ForeColor = CivicCalmTheme.ColorInk}
            cmbDeadlinePreset.Items.AddRange(New Object() {"24 Hours", "3 Days", "7 Days", "15 Days", "Custom"})
            cmbDeadlinePreset.SelectedIndex = 1

            ' The picker only applies on the Custom preset; UpdateDeadlineState keeps it
            ' disabled (and the hint clear) until Custom is selected.
            dtpDeadline = New DateTimePicker With {
                .Format = DateTimePickerFormat.Custom,
                .CustomFormat = "yyyy-MM-dd HH:mm",
                .Height = 28,
                .TabIndex = 5,
                .Enabled = False,
                .Value = DateTime.Now.AddDays(3),
                .BackColor = CivicCalmTheme.ColorSurface,
                .ForeColor = CivicCalmTheme.ColorInk
            }
            lblDeadlineHint = New Label With {
                .Text = "",
                .AutoSize = True,
                .Font = CivicCalmTheme.FontMicrocopy,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .Margin = New Padding(0, 0, 0, 6)
            }
            AddHandler cmbDeadlinePreset.SelectedIndexChanged, Sub()
                                                                   If cmbDeadlinePreset.SelectedItem IsNot Nothing Then
                                                                       Select Case cmbDeadlinePreset.SelectedItem.ToString()
                                                                           Case "24 Hours" : dtpDeadline.Value = DateTime.Now.AddHours(24)
                                                                           Case "3 Days" : dtpDeadline.Value = DateTime.Now.AddDays(3)
                                                                           Case "7 Days" : dtpDeadline.Value = DateTime.Now.AddDays(7)
                                                                           Case "15 Days" : dtpDeadline.Value = DateTime.Now.AddDays(15)
                                                                       End Select
                                                                   End If
                                                                   UpdateDeadlineState()
                                                               End Sub
            AddHandler dtpDeadline.ValueChanged, Sub() UpdateDeadlineState()
            UpdateDeadlineState()

            cmbOrigin = NewSuggestionCombo(6)
            cmbDest = NewSuggestionCombo(7)
            cmbCabinet = NewSuggestionCombo(8)
            cmbShelf = NewSuggestionCombo(9)
            cmbBox = NewSuggestionCombo(10)

            txtGDrive = New TextBox With {.Height = 28, .Text = "", .TabIndex = 11, .BackColor = CivicCalmTheme.ColorSurface, .ForeColor = CivicCalmTheme.ColorInk, .BorderStyle = BorderStyle.FixedSingle}

            btnAttachScan = New Button With {
                .Text = "&Attach / Scan Document",
                .Size = New Size(340, 32),
                .TabIndex = 12,
                .BackColor = CivicCalmTheme.ColorWell,
                .ForeColor = CivicCalmTheme.ColorInk,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontMicrocopy,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(0, 2, 0, 6)
            }
            btnAttachScan.FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder
            AddHandler btnAttachScan.Click, AddressOf OnAttachScanDocument

            cmbAssignedStaff = New ComboBox With {.DropDownStyle = ComboBoxStyle.DropDownList, .Height = 28, .TabIndex = 13, .BackColor = CivicCalmTheme.ColorSurface, .ForeColor = CivicCalmTheme.ColorInk}

            btnRegister = New Button With {
                .Text = "&Register OSG Document",
                .Height = 42,
                .TabIndex = 14,
                .BackColor = CivicCalmTheme.ColorPrimary,
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(0, 12, 0, 12)
            }
            btnRegister.FlatAppearance.BorderSize = 0
            AddHandler btnRegister.Click, AddressOf OnRegisterDocument

            ' Six equal columns let related fields share a row: a full-width field spans six, a
            ' pair spans three, the landmark triple spans two each. One field per row made the
            ' form taller than its card, so the desk scrolled a form that already had room.
            Dim pnlFormFlow As New TableLayoutPanel With {
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .ColumnCount = 6,
                .BackColor = CivicCalmTheme.ColorSurface
            }
            For c As Integer = 0 To 5
                pnlFormFlow.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F / 6.0F))
            Next

            Dim nextRow As Integer = 0
            Dim openLabel As Integer = -1
            Dim colsUsed As Integer = 0

            Dim EnsureOpen = Sub()
                                  If openLabel >= 0 Then Return
                                  openLabel = nextRow
                                  colsUsed = 0
                                  pnlFormFlow.RowStyles.Add(New RowStyle(SizeType.AutoSize))
                                  pnlFormFlow.RowStyles.Add(New RowStyle(SizeType.AutoSize))
                                  nextRow += 2
                              End Sub

            Dim AddField = Sub(labelText As String, ctrl As Control, span As Integer)
                                EnsureOpen()
                                Dim lbl As New Label With {
                                    .Text = labelText,
                                    .Dock = DockStyle.Fill,
                                    .Font = CivicCalmTheme.FontFieldLabel,
                                    .ForeColor = CivicCalmTheme.ColorInkMuted,
                                    .TextAlign = ContentAlignment.BottomLeft,
                                    .Margin = New Padding(0, 6, 8, 0)
                                }
                                ctrl.Dock = DockStyle.Fill
                                ctrl.Margin = New Padding(0, 0, 8, 8)
                                pnlFormFlow.Controls.Add(lbl, colsUsed, openLabel)
                                pnlFormFlow.SetColumnSpan(lbl, span)
                                pnlFormFlow.Controls.Add(ctrl, colsUsed, openLabel + 1)
                                pnlFormFlow.SetColumnSpan(ctrl, span)
                                colsUsed += span
                                If colsUsed >= 6 Then openLabel = -1
                            End Sub

            ' A control with no label of its own: it takes a whole row and closes any row still
            ' open, so a wide item never lands beside a field's leftover column space.
            Dim AddWide = Sub(ctrl As Control)
                               openLabel = -1
                               EnsureOpen()
                               ctrl.Dock = DockStyle.Fill
                               pnlFormFlow.Controls.Add(ctrl, 0, openLabel)
                               pnlFormFlow.SetColumnSpan(ctrl, 6)
                               openLabel = -1
                           End Sub

            AddField("Document Title / Subject:", txtTitle, 6)
            AddField("Classification Category:", cmbDocType, 3)
            AddField("Flow Direction:", cmbFlowDirection, 3)
            AddField("Processing Deadline Preset:", cmbDeadlinePreset, 3)
            AddField("Target Deadline (Date/Time):", dtpDeadline, 3)
            AddWide(lblDeadlineHint)
            AddField("Originating Office:", cmbOrigin, 3)
            AddField("Destination Office:", cmbDest, 3)
            AddField("Cabinet ID:", cmbCabinet, 2)
            AddField("Shelf No:", cmbShelf, 2)
            AddField("Box Code:", cmbBox, 2)
            AddField("Digital Soft Copy Reference:", txtGDrive, 6)
            AddWide(btnAttachScan)
            AddField("Assigned Staff Officer:", cmbAssignedStaff, 6)
            AddWide(btnRegister)

            pnlFormScroll.Controls.Add(pnlFormFlow)
            pnlFormCard.Controls.Add(pnlFormScroll)

            ' Right: Grid Card Panel
            Dim pnlGridCard As New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(12)
            }
            ApplyCardBorder(pnlGridCard)

            Dim flwToolbar As New FlowLayoutPanel With {
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .FlowDirection = FlowDirection.LeftToRight,
                .WrapContents = True,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(0, 0, 0, 8)
            }
            btnViewRegistryDetail = New Button With {
                .Text = "&View Details && History",
                .Size = New Size(220, 34),
                .BackColor = CivicCalmTheme.ColorWell,
                .ForeColor = CivicCalmTheme.ColorInk,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(0, 0, 8, 0)
            }
            btnViewRegistryDetail.FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder
            AddHandler btnViewRegistryDetail.Click, Sub() OpenSelectedDocumentDetail(dgvRegistry)
            flwToolbar.Controls.Add(btnViewRegistryDetail)

            Dim btnPrintRegistrySlip As New Button With {
                .Text = "&Print Routing Slip",
                .Size = New Size(180, 34),
                .BackColor = CivicCalmTheme.ColorWell,
                .ForeColor = CivicCalmTheme.ColorInk,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(0, 0, 0, 0)
            }
            btnPrintRegistrySlip.FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder
            AddHandler btnPrintRegistrySlip.Click, Sub() PrintSelectedDocumentRoutingSlip(dgvRegistry)
            flwToolbar.Controls.Add(btnPrintRegistrySlip)

            dgvRegistry = New DataGridView With {
                .Dock = DockStyle.Fill,
                .ReadOnly = True,
                .AllowUserToAddRows = False
            }
            DataGridStyler.ApplyCivicStyle(dgvRegistry)
            AddHandler dgvRegistry.CellDoubleClick, Sub(s, e) If e.RowIndex >= 0 Then OpenSelectedDocumentDetail(dgvRegistry)

            lblRegistryWatermark = CreateGridWatermark(pnlGridCard)

            pnlGridCard.Controls.Add(dgvRegistry)
            pnlGridCard.Controls.Add(flwToolbar)

            tblMain.Controls.Add(pnlFormCard, 0, 0)
            tblMain.Controls.Add(pnlGridCard, 1, 0)
            viewRegistry.Controls.Add(tblMain)

            LoadRegistrySuggestions()
        End Sub

        ' Editable dropdown that offers values already used on earlier documents: the arrow
        ' lists the full history, typing filters it, and free text is still accepted.
        Private Function NewSuggestionCombo(tabIndex As Integer) As ComboBox
            Dim cb As New ComboBox With {
                .DropDownStyle = ComboBoxStyle.DropDown,
                .Height = 28,
                .TabIndex = tabIndex,
                .DropDownWidth = 460,
                .BackColor = CivicCalmTheme.ColorSurface,
                .ForeColor = CivicCalmTheme.ColorInk
            }
            Try
                If System.Threading.Thread.CurrentThread.GetApartmentState() = System.Threading.ApartmentState.STA Then
                    cb.AutoCompleteMode = AutoCompleteMode.SuggestAppend
                    cb.AutoCompleteSource = AutoCompleteSource.ListItems
                End If
            Catch ex As System.Threading.ThreadStateException
                ' Suppress in MTA test runners where OLE COM autocomplete is unsupported
            End Try
            Return cb
        End Function

        ' Seed the registry dropdowns with every value already on file so clerks repick
        ' instead of retyping. The document rows are the only store: a registered value
        ' persists with its document, and sync brings in rows from other workstations.
        Private Sub LoadRegistrySuggestions()
            If cmbOrigin Is Nothing Then Return
            FillSuggestionCombo(cmbOrigin, DistinctDocumentColumn("OriginatingOffice"))
            FillSuggestionCombo(cmbDest, DistinctDocumentColumn("DestinationOffice"))
            FillSuggestionCombo(cmbCabinet, DistinctDocumentColumn("CabinetID"))
            FillSuggestionCombo(cmbShelf, DistinctDocumentColumn("ShelfNo"))
            FillSuggestionCombo(cmbBox, DistinctDocumentColumn("BoxCode"))
        End Sub

        Private Sub FillSuggestionCombo(cmb As ComboBox, values As List(Of String))
            Dim typed As String = cmb.Text
            cmb.BeginUpdate()
            Try
                cmb.Items.Clear()
                For Each v In values
                    cmb.Items.Add(v)
                Next
                cmb.Text = typed
            Finally
                cmb.EndUpdate()
            End Try
        End Sub

        Private Function DistinctDocumentColumn(columnName As String) As List(Of String)
            Dim values As New List(Of String)()
            Dim dt = EmbeddedDB.DataSet.Tables("Documents")
            If dt Is Nothing OrElse Not dt.Columns.Contains(columnName) Then Return values
            Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each row As DataRow In dt.Rows
                If IsDBNull(row(columnName)) Then Continue For
                Dim v = row(columnName).ToString().Trim()
                If v.Length > 0 AndAlso seen.Add(v) Then values.Add(v)
            Next
            values.Sort(StringComparer.CurrentCultureIgnoreCase)
            Return values
        End Function

        ' The picker only applies on the Custom preset; presets compute the deadline
        ' themselves. Custom values get a live hint: amber when the deadline lands within
        ' 24 hours, red plus a blocking error when it is already past.
        Private Sub UpdateDeadlineState()
            Dim isCustom As Boolean = cmbDeadlinePreset.SelectedItem IsNot Nothing AndAlso cmbDeadlinePreset.SelectedItem.ToString().Equals("Custom", StringComparison.OrdinalIgnoreCase)
            dtpDeadline.Enabled = isCustom
            epValidation.SetError(dtpDeadline, "")
            If Not isCustom Then
                lblDeadlineHint.Text = ""
                Return
            End If
            If dtpDeadline.Value <= DateTime.Now Then
                lblDeadlineHint.Text = "Target deadline cannot be in the past."
                lblDeadlineHint.ForeColor = CivicCalmTheme.ColorDanger
                epValidation.SetError(dtpDeadline, "Pick a date and time later than now, or choose a preset.")
            ElseIf dtpDeadline.Value <= DateTime.Now.AddHours(24) Then
                lblDeadlineHint.Text = "This deadline falls due within 24 hours."
                lblDeadlineHint.ForeColor = CivicCalmTheme.ColorWarning
            Else
                lblDeadlineHint.Text = ""
            End If
        End Sub

        Private Function DeadlineValidationError() As String
            If cmbDeadlinePreset.SelectedItem Is Nothing OrElse Not cmbDeadlinePreset.SelectedItem.ToString().Equals("Custom", StringComparison.OrdinalIgnoreCase) Then Return ""
            If dtpDeadline.Value <= DateTime.Now Then Return "Target deadline cannot be in the past. Pick a later date and time."
            Return ""
        End Function

        Private Sub OnAttachScanDocument(sender As Object, e As EventArgs)
            Using ofd As New OpenFileDialog()
                ofd.Title = "Select Scanned Document or PDF Attachment"
                ' PDF only: the launch and preview paths (ValidateGDriveURL) accept local .pdf
                ' files, so inviting the operator to attach a JPG or DOCX would store a link
                ' the system then refuses to open.
                ofd.Filter = "PDF Files (*.pdf)|*.pdf"
                If ofd.ShowDialog(Me) = DialogResult.OK Then
                    txtGDrive.Text = ofd.FileName
                    lblStatusMessage.Text = "Attachment selected: " & System.IO.Path.GetFileName(ofd.FileName)
                End If
            End Using
        End Sub

        Private Function ValidateRegistrationInputs() As Boolean
            If CurrentUser Is Nothing Then
                lblStatusMessage.Text = "Authentication Required: Tap RFID card to register documents."
                MessageBox.Show("RFID Authentication Required to Register Documents.", "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return False
            End If

            If String.IsNullOrWhiteSpace(txtTitle.Text) Then
                epValidation.SetError(txtTitle, "Please enter Document Title.")
                lblStatusMessage.Text = "Validation Error: Document Title is required."
                txtTitle.Focus()
                Return False
            End If

            Dim deadlineError As String = DeadlineValidationError()
            If deadlineError <> "" Then
                epValidation.SetError(dtpDeadline, deadlineError)
                lblDeadlineHint.Text = "Target deadline cannot be in the past."
                lblDeadlineHint.ForeColor = CivicCalmTheme.ColorDanger
                lblStatusMessage.Text = "Validation Error: " & deadlineError
                dtpDeadline.Focus()
                Return False
            End If

            Dim errUrlMsg As String = ""
            If Not String.IsNullOrWhiteSpace(txtGDrive.Text) AndAlso txtGDrive.Text.StartsWith("http", StringComparison.OrdinalIgnoreCase) Then
                If Not EmbeddedDB.ValidateGDriveURL(txtGDrive.Text, errUrlMsg) Then
                    epValidation.SetError(txtGDrive, errUrlMsg)
                    lblStatusMessage.Text = "Invalid Soft Copy URL: " & errUrlMsg
                    txtGDrive.Focus()
                    Return False
                End If
            End If

            Return True
        End Function

        Private Sub ClearRegistrationInputs()
            txtTitle.Text = ""
            cmbOrigin.Text = ""
            cmbDest.Text = ""
            cmbCabinet.Text = ""
            cmbShelf.Text = ""
            cmbBox.Text = ""
            txtGDrive.Text = ""
            If cmbDocType.Items.Count > 0 Then cmbDocType.SelectedIndex = 0
            If cmbFlowDirection.Items.Count > 0 Then cmbFlowDirection.SelectedIndex = 0
            If cmbDeadlinePreset.Items.Count > 1 Then cmbDeadlinePreset.SelectedIndex = 1
            epValidation.Clear()
            txtTitle.Focus()
        End Sub

        Private Sub OnRegisterDocument(sender As Object, e As EventArgs)
            epValidation.Clear()
            If Not ValidateRegistrationInputs() Then Return

            Dim docType = If(cmbDocType.SelectedItem IsNot Nothing, cmbDocType.SelectedItem.ToString(), "Regular Communication")
            Dim flowDir = If(cmbFlowDirection.SelectedItem IsNot Nothing, cmbFlowDirection.SelectedItem.ToString().ToUpperInvariant(), "INCOMING")
            Dim code = EmbeddedDB.GenerateDocCode(docType)
            Dim assigned = If(cmbAssignedStaff.SelectedItem IsNot Nothing, cmbAssignedStaff.SelectedItem.ToString(), "")
            Dim assignedSec = DocumentService.GetDefaultSectionForCategory(docType)
            Dim targetDeadline = dtpDeadline.Value.ToString("yyyy-MM-dd HH:mm:ss")
            Dim initialStatus = If(docType.Equals("Finance", StringComparison.OrdinalIgnoreCase), "RECEIVED", "FOR_REVIEW")
            Dim lastAction = "Registered by Records Section and auto-routed to " & assignedSec

            Dim currentUserId As Integer = 1
            If CurrentUser IsNot Nothing AndAlso CurrentUser.Table.Columns.Contains("UserID") AndAlso Not IsDBNull(CurrentUser("UserID")) Then
                currentUserId = Convert.ToInt32(CurrentUser("UserID"))
            End If
            Dim currentUserName As String = If(CurrentUser IsNot Nothing, CurrentUser("FullName").ToString(), "System Staff")

            If Program.Coordinator IsNot Nothing Then
                ' The coordinator mints the code in SQL when it can reach the server, so the
                ' code that comes back is the one that was filed. The local code above is only
                ' the offline fallback.
                code = Program.Coordinator.RegisterDocument(code, docType, txtTitle.Text.Trim(), cmbOrigin.Text.Trim(), cmbDest.Text.Trim(), cmbCabinet.Text.Trim(), cmbShelf.Text.Trim(), cmbBox.Text.Trim(), txtGDrive.Text.Trim(), initialStatus, assigned, flowDir, assignedSec, targetDeadline, "", lastAction, currentUserName, currentUserId)
            Else
                EmbeddedDB.AddDocument(code, docType, txtTitle.Text.Trim(), cmbOrigin.Text.Trim(), cmbDest.Text.Trim(), cmbCabinet.Text.Trim(), cmbShelf.Text.Trim(), cmbBox.Text.Trim(), txtGDrive.Text.Trim(), initialStatus, assigned, flowDir, assignedSec, targetDeadline, "", lastAction)
                EmbeddedDB.LogAudit(currentUserName, String.Format("Registered New OSG Document [{0}] : {1} (Auto-routed to {2})", code, txtTitle.Text.Trim(), assignedSec))
            End If

            lblStatusMessage.Text = String.Format("Document Registered: {0} [{1}] routed to {2}", code, txtTitle.Text.Trim(), assignedSec)
            ClearRegistrationInputs()
            RefreshActiveTabGrid()
        End Sub
    End Class
End Namespace
