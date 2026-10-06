Option Explicit On
Option Strict On

Imports System
Imports System.Data
Imports System.Drawing
Imports System.Windows.Forms

Namespace BTA_OSG
    Partial Public Class FormMain
        Private Sub SetupAdminView()
            viewAdmin = New Panel With {.Dock = DockStyle.Fill, .BackColor = CivicCalmTheme.ColorCanvas}

            Dim tblMain As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 2,
                .RowCount = 2
            }
            tblMain.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, Dpi(380.0F)))
            tblMain.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            tblMain.RowStyles.Add(New RowStyle(SizeType.Percent, 58.0F))
            tblMain.RowStyles.Add(New RowStyle(SizeType.Percent, 42.0F))

            Dim pnlFormCard As New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(16),
                .Margin = New Padding(0, 0, 16, 0)
            }
            ApplyCardBorder(pnlFormCard)

            Dim pnlFormFlow As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.TopDown,
                .WrapContents = False,
                .AutoScroll = True,
                .BackColor = CivicCalmTheme.ColorSurface
            }

            Dim lblHeader As New Label With {
                .Text = "REGISTER USER & RFID BADGE",
                .UseMnemonic = False,
                .Font = CivicCalmTheme.FontSectionHeader,
                .ForeColor = CivicCalmTheme.ColorInk,
                .AutoSize = True,
                .Margin = New Padding(0, 0, 0, 12)
            }

            txtNewUserName = New TextBox With {.Width = CInt(Dpi(320.0F)), .Height = CInt(Dpi(28.0F)), .Text = "", .TabIndex = 1, .BackColor = CivicCalmTheme.ColorSurface, .ForeColor = CivicCalmTheme.ColorInk, .BorderStyle = BorderStyle.FixedSingle}

            cmbNewUserRole = New ComboBox With {.Width = CInt(Dpi(320.0F)), .Height = CInt(Dpi(28.0F)), .TabIndex = 2, .DropDownStyle = ComboBoxStyle.DropDownList, .BackColor = CivicCalmTheme.ColorSurface, .ForeColor = CivicCalmTheme.ColorInk}
            cmbNewUserRole.Items.AddRange(New Object() {
                "Administrative Staff",
                "Section Officer / Clerk",
                "Records Section",
                "Secretariat",
                "Legislative Section",
                "Finance Section",
                "Travel Section",
                "Secretary-General",
                "OSG Chief",
                "System Administrator"
            })
            cmbNewUserRole.SelectedIndex = 0
            AddHandler cmbNewUserRole.SelectedIndexChanged, AddressOf OnRoleSelectionChanged

            cmbNewUserSection = New ComboBox With {.Width = CInt(Dpi(320.0F)), .Height = CInt(Dpi(28.0F)), .TabIndex = 3, .DropDownStyle = ComboBoxStyle.DropDownList, .BackColor = CivicCalmTheme.ColorSurface, .ForeColor = CivicCalmTheme.ColorInk}
            cmbNewUserSection.Items.AddRange(New Object() {
                "Records Section",
                "Secretariat",
                "Legislative Section",
                "Finance Section",
                "Travel Section",
                "General Operations / Front Desk",
                "Office of the Secretary-General",
                "ICT / Systems Administration"
            })
            cmbNewUserSection.SelectedIndex = 0
            AddHandler cmbNewUserSection.SelectedIndexChanged, Sub() UpdateSectionNotice()

            lblSectionNotice = New Label With {
                .Width = CInt(Dpi(320.0F)),
                .Font = CivicCalmTheme.FontMicrocopy,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .AutoSize = False,
                .Height = CInt(Dpi(30.0F)),
                .AutoEllipsis = True,
                .Margin = New Padding(0, 2, 0, 4),
                .Text = "Stationed at Records: Handles receiving, digitizing intake, and digital release."
            }

            txtNewUserUID = New TextBox With {.Width = CInt(Dpi(320.0F)), .Height = CInt(Dpi(28.0F)), .Text = "", .TabIndex = 4, .BackColor = CivicCalmTheme.ColorSurface, .ForeColor = CivicCalmTheme.ColorInk, .BorderStyle = BorderStyle.FixedSingle}

            chkCanRoute = New CheckBox With {
                .Text = "Allow &Cross-Section Routing",
                .Font = CivicCalmTheme.FontBody,
                .ForeColor = CivicCalmTheme.ColorInk,
                .AutoSize = True,
                .Checked = False,
                .TabIndex = 5,
                .Margin = New Padding(0, 2, 0, 2)
            }

            chkCanMove = New CheckBox With {
                .Text = "Allow Storage Landmark &Transfers",
                .Font = CivicCalmTheme.FontBody,
                .ForeColor = CivicCalmTheme.ColorInk,
                .AutoSize = True,
                .Checked = True,
                .TabIndex = 6,
                .Margin = New Padding(0, 2, 0, 2)
            }

            chkCanSoftCopy = New CheckBox With {
                .Text = "Allow S&oft Copy Document Access",
                .Font = CivicCalmTheme.FontBody,
                .ForeColor = CivicCalmTheme.ColorInk,
                .AutoSize = True,
                .Checked = True,
                .TabIndex = 7,
                .Margin = New Padding(0, 2, 0, 2)
            }

            btnAddUser = New Button With {
                .Text = " Sa&ve User && RFID Smart Card",
                .Size = New Size(320, 42),
                .TabIndex = 8,
                .BackColor = CivicCalmTheme.ColorPrimary,
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(0, 14, 0, 0),
                .Image = AppAssets.GetIcon("floppy-disk", 18, Color.White),
                .ImageAlign = ContentAlignment.MiddleLeft,
                .TextAlign = ContentAlignment.MiddleCenter,
                .TextImageRelation = TextImageRelation.ImageBeforeText,
                .Padding = New Padding(8, 0, 8, 0)
            }
            btnAddUser.FlatAppearance.BorderSize = 0
            AddHandler btnAddUser.Click, AddressOf OnAddUser

            ' The in-app unlock the office docs describe: a badge that failed its tap
            ' threshold stays locked in tbl_Users across restarts, so without this the
            ' remedy is hand-editing SQL on the server.
            btnUnlockUser = New Button With {
                .Text = " &Unlock Selected Account",
                .Size = New Size(320, 38),
                .TabIndex = 9,
                .BackColor = CivicCalmTheme.ColorWell,
                .ForeColor = CivicCalmTheme.ColorInk,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontBody,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(0, 8, 0, 0),
                .Image = AppAssets.GetIcon("lock-key", 16, CivicCalmTheme.ColorInk),
                .ImageAlign = ContentAlignment.MiddleLeft,
                .TextAlign = ContentAlignment.MiddleCenter,
                .TextImageRelation = TextImageRelation.ImageBeforeText,
                .Padding = New Padding(8, 0, 8, 0)
            }
            btnUnlockUser.FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder
            AddHandler btnUnlockUser.Click, AddressOf OnUnlockSelectedUser

            ' Lost-badge recovery: revocation used to exist only as a service method with no
            ' UI, so a badge reported lost stayed valid everywhere until someone edited SQL.
            btnRevokeCard = New Button With {
                .Text = " Revoke Selected &Badge",
                .Size = New Size(320, 38),
                .TabIndex = 10,
                .BackColor = CivicCalmTheme.ColorWell,
                .ForeColor = CivicCalmTheme.ColorInk,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontBody,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(0, 8, 0, 0),
                .Image = AppAssets.GetIcon("x", 16, CivicCalmTheme.ColorInk),
                .ImageAlign = ContentAlignment.MiddleLeft,
                .TextAlign = ContentAlignment.MiddleCenter,
                .TextImageRelation = TextImageRelation.ImageBeforeText,
                .Padding = New Padding(8, 0, 8, 0)
            }
            btnRevokeCard.FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder
            AddHandler btnRevokeCard.Click, AddressOf OnRevokeSelectedCard

            pnlFormFlow.Controls.Add(lblHeader)
            pnlFormFlow.Controls.Add(New Label With {.Text = "&Full Name:", .Font = CivicCalmTheme.FontFieldLabel, .ForeColor = CivicCalmTheme.ColorInkMuted, .AutoSize = True, .Margin = New Padding(0, 6, 0, 2)})
            pnlFormFlow.Controls.Add(txtNewUserName)
            pnlFormFlow.Controls.Add(New Label With {.Text = "Syste&m Role:", .Font = CivicCalmTheme.FontFieldLabel, .ForeColor = CivicCalmTheme.ColorInkMuted, .AutoSize = True, .Margin = New Padding(0, 6, 0, 2)})
            pnlFormFlow.Controls.Add(cmbNewUserRole)
            pnlFormFlow.Controls.Add(New Label With {.Text = "Assigned &Section Desk / Department:", .Font = CivicCalmTheme.FontFieldLabel, .ForeColor = CivicCalmTheme.ColorInkMuted, .AutoSize = True, .Margin = New Padding(0, 6, 0, 2)})
            pnlFormFlow.Controls.Add(cmbNewUserSection)
            pnlFormFlow.Controls.Add(lblSectionNotice)
            pnlFormFlow.Controls.Add(New Label With {.Text = "&RFID Smart Card UID (Hex):", .Font = CivicCalmTheme.FontFieldLabel, .ForeColor = CivicCalmTheme.ColorInkMuted, .AutoSize = True, .Margin = New Padding(0, 6, 0, 2)})
            pnlFormFlow.Controls.Add(txtNewUserUID)
            pnlFormFlow.Controls.Add(New Label With {.Text = "Staff Operational Privileges:", .Font = CivicCalmTheme.FontFieldLabel, .ForeColor = CivicCalmTheme.ColorInkMuted, .AutoSize = True, .Margin = New Padding(0, 8, 0, 2)})
            pnlFormFlow.Controls.Add(chkCanRoute)
            pnlFormFlow.Controls.Add(chkCanMove)
            pnlFormFlow.Controls.Add(chkCanSoftCopy)
            pnlFormFlow.Controls.Add(btnAddUser)
            pnlFormFlow.Controls.Add(btnUnlockUser)
            pnlFormFlow.Controls.Add(btnRevokeCard)

            Dim btnSetupWizard As New Button With {
                .Text = " Reconfigure Station Set&up...",
                .Size = New Size(CInt(Dpi(320.0F)), CInt(Dpi(38.0F))),
                .TabIndex = 11,
                .BackColor = CivicCalmTheme.ColorWell,
                .ForeColor = CivicCalmTheme.ColorInk,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontBody,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(0, 12, 0, 0),
                .Image = AppAssets.GetIcon("gear", 16, CivicCalmTheme.ColorInk),
                .ImageAlign = ContentAlignment.MiddleLeft,
                .TextAlign = ContentAlignment.MiddleCenter,
                .TextImageRelation = TextImageRelation.ImageBeforeText,
                .Padding = New Padding(8, 0, 8, 0)
            }
            btnSetupWizard.FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder
            AddHandler btnSetupWizard.Click, Async Sub()
                                                If Not FormConnect.AuthorizeConfiguredChange() Then Return
                                                Using connect As New FormConnect()
                                                    If connect.ShowDialog(Me) <> DialogResult.OK Then Return
                                                    ' FormConnect already confirmed what changed; a second
                                                    ' dialog here only repeats it. The reload has to happen in
                                                    ' this session or the station reads one server and writes
                                                    ' another.
                                                    Program.ApplyStationSetup()
                                                    Await RunSqlSyncAsync()
                                                    RefreshActiveTabGrid()
                                                    lblStatusMessage.Text = If(Program.IsDatabaseConnected,
                                                        "Station setup applied: office mode, reading and writing the office database.",
                                                        "Station setup applied: demo mode, working from the local cache.")
                                                End Using
                                            End Sub
            pnlFormFlow.Controls.Add(btnSetupWizard)

            pnlFormCard.Controls.Add(pnlFormFlow)

            Dim pnlGridCard As New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(12)
            }
            ApplyCardBorder(pnlGridCard)

            dgvUsers = New DataGridView With {
                .Dock = DockStyle.Fill,
                .ReadOnly = True,
                .AllowUserToAddRows = False
            }
            DataGridStyler.ApplyCivicStyle(dgvUsers)

            lblUsersWatermark = CreateGridWatermark(pnlGridCard)

            pnlGridCard.Controls.Add(dgvUsers)

            Dim pnlSeatsCard As New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(12)
            }
            ApplyCardBorder(pnlSeatsCard)

            Dim lblSeatsHeader As New Label With {
                .Text = "WORKSTATION SYNC STATUS",
                .UseMnemonic = False,
                .Font = CivicCalmTheme.FontSectionHeader,
                .ForeColor = CivicCalmTheme.ColorInk,
                .Dock = DockStyle.Top,
                .AutoSize = True
            }

            dgvSeats = New DataGridView With {
                .Dock = DockStyle.Fill,
                .ReadOnly = True,
                .AllowUserToAddRows = False
            }
            DataGridStyler.ApplyCivicStyle(dgvSeats)

            lblSeatsWatermark = CreateGridWatermark(pnlSeatsCard)

            pnlSeatsCard.Controls.Add(dgvSeats)
            pnlSeatsCard.Controls.Add(lblSeatsHeader)

            tblMain.Controls.Add(pnlFormCard, 0, 0)
            tblMain.SetRowSpan(pnlFormCard, 2)
            tblMain.Controls.Add(pnlGridCard, 1, 0)
            tblMain.Controls.Add(pnlSeatsCard, 1, 1)
            viewAdmin.Controls.Add(tblMain)
        End Sub

        Private Sub OnRoleSelectionChanged(sender As Object, e As EventArgs)
            If cmbNewUserRole.SelectedItem Is Nothing Then Return
            Dim selectedRole = cmbNewUserRole.SelectedItem.ToString()

            Select Case selectedRole
                Case "Administrative Staff"
                    ApplyRolePreset("", "", canRoute:=False, isSectionFixed:=False)

                Case "Section Officer / Clerk"
                    ApplyRolePreset("", "", canRoute:=True, isSectionFixed:=False)

                Case "Records Section"
                    ApplyRolePreset("Records Section", "Records Desk Officer: Authorized for intake, indexing, and official release.", canRoute:=True)

                Case "Secretariat"
                    ApplyRolePreset("Secretariat", "Secretariat Officer: Regular communications intake and executive review.", canRoute:=True)

                Case "Legislative Section"
                    ApplyRolePreset("Legislative Section", "Legislative Desk Officer: Committee reports, bill drafts, and resolution tracking.", canRoute:=True)

                Case "Finance Section"
                    ApplyRolePreset("Finance Section", "Finance Desk Officer: Operational budgets, disbursement vouchers, and audit compliance.", canRoute:=True)

                Case "Travel Section"
                    ApplyRolePreset("Travel Section", "Travel Desk Officer: Official mission travel orders, itineraries, and allowances.", canRoute:=True)

                Case "Secretary-General", "OSG Chief"
                    ApplyRolePreset("Office of the Secretary-General", "Executive Leadership: Global document visibility, directives, and approvals.", canRoute:=True)

                Case "System Administrator"
                    ApplyRolePreset("ICT / Systems Administration", "System Administrator: User management, RFID provisioning, and audit logs.", canRoute:=True)
            End Select
        End Sub

        Private Sub ApplyRolePreset(sectionName As String, noticeText As String, canRoute As Boolean, Optional isSectionFixed As Boolean = True)
            If isSectionFixed AndAlso Not String.IsNullOrEmpty(sectionName) Then
                Dim idx = cmbNewUserSection.Items.IndexOf(sectionName)
                If idx >= 0 Then cmbNewUserSection.SelectedIndex = idx
                cmbNewUserSection.Enabled = False
                lblSectionNotice.Text = noticeText
            Else
                cmbNewUserSection.Enabled = True
                UpdateSectionNotice()
            End If

            chkCanRoute.Checked = canRoute
            chkCanMove.Checked = True
            chkCanSoftCopy.Checked = True
        End Sub

        Private Sub UpdateSectionNotice()
            If cmbNewUserSection.SelectedItem Is Nothing Then Return
            Dim sec = cmbNewUserSection.SelectedItem.ToString()
            Select Case sec
                Case "Records Section"
                    lblSectionNotice.Text = "Stationed at Records: Handles receiving, digitizing intake, and digital release."
                Case "Secretariat"
                    lblSectionNotice.Text = "Stationed at Secretariat: Reviews regular communications and executive endorsements."
                Case "Legislative Section"
                    lblSectionNotice.Text = "Stationed at Legislative: Manages Parliament bills, resolutions, and committee reports."
                Case "Finance Section"
                    lblSectionNotice.Text = "Stationed at Finance: Processes budget allocations, vouchers, and cash flow tracking."
                Case "Travel Section"
                    lblSectionNotice.Text = "Stationed at Travel: Coordinates official mission orders and travel authorizations."
                Case "General Operations / Front Desk"
                    lblSectionNotice.Text = "Stationed at Front Desk: General administrative support and visitor intake."
                Case "Office of the Secretary-General"
                    lblSectionNotice.Text = "Stationed at OSG Executive Office: Direct executive assistance."
                Case "ICT / Systems Administration"
                    lblSectionNotice.Text = "Stationed at ICT / SysAdmin: Technical systems and database administration."
            End Select
        End Sub

        Private Sub OnAddUser(sender As Object, e As EventArgs)
            epValidation.Clear()
            If CurrentUser Is Nothing OrElse CurrentUser("Role").ToString() <> "System Administrator" Then
                lblStatusMessage.Text = "Access Denied: System Administrator privileges required."
                MessageBox.Show("System Administrator Privileges Required.", "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim cleanName = txtNewUserName.Text.Trim()
            ' The Admin box accepts pasted reader dumps, so it goes through the same
            ' control-character sanitizer as the login and claim paths: a UID stored with a
            ' trailing control character enrols fine but can never authenticate.
            Dim cleanUid = EmbeddedDB.SanitizeCardUid(txtNewUserUID.Text)

            If String.IsNullOrWhiteSpace(cleanName) Then
                epValidation.SetError(txtNewUserName, "Please enter Full Name.")
                lblStatusMessage.Text = "Validation Error: Full Name is required."
                txtNewUserName.Focus()
                Return
            End If

            If String.IsNullOrWhiteSpace(cleanUid) Then
                epValidation.SetError(txtNewUserUID, "Please enter RFID Card UID.")
                lblStatusMessage.Text = "Validation Error: RFID Card UID is required."
                txtNewUserUID.Focus()
                Return
            End If

            Dim selectedRole = If(cmbNewUserRole.SelectedItem IsNot Nothing, cmbNewUserRole.SelectedItem.ToString(), "Administrative Staff")
            Dim selectedSection = If(cmbNewUserSection.SelectedItem IsNot Nothing, cmbNewUserSection.SelectedItem.ToString(), "Records Section")

            Dim savedRole As String
            Dim savedOffice As String

            If selectedRole = "Administrative Staff" Then
                savedRole = "Administrative Staff"
                savedOffice = selectedSection
            ElseIf selectedRole = "Section Officer / Clerk" Then
                savedRole = selectedSection
                savedOffice = selectedSection
            Else
                savedRole = selectedRole
                savedOffice = selectedSection
            End If

            Dim maskedUid = If(cleanUid.Length > 4, "****" & cleanUid.Substring(cleanUid.Length - 4), cleanUid)
            If Program.Coordinator IsNot Nothing Then
                ' Connected enrolment lands in tbl_Users, tbl_RfidCards, and tbl_UserRoles, so
                ' the badge works on every workstation, not just this one.
                Dim enrolledBy = CurrentUser("FullName").ToString()
                Dim enrolledById As Integer = 1
                If CurrentUser.Table.Columns.Contains("UserID") AndAlso Not IsDBNull(CurrentUser("UserID")) Then
                    enrolledById = Convert.ToInt32(CurrentUser("UserID"))
                End If
                Program.Coordinator.RegisterUser(cleanName, savedRole, savedOffice, cleanUid, chkCanRoute.Checked, chkCanMove.Checked, chkCanSoftCopy.Checked, enrolledBy, enrolledById)
            Else
                ' pendingSync:=True: the coordinator-less fallback must still replay, like
                ' every other offline mutator.
                EmbeddedDB.AddUser(cleanUid, cleanName, savedRole, savedOffice, chkCanRoute.Checked, chkCanMove.Checked, chkCanSoftCopy.Checked, pendingSync:=True)
                EmbeddedDB.LogAudit(CurrentUser("FullName").ToString(), String.Format("Registered/Updated User [{0}] Role: {1} (Desk: {2}) Privileges: [Route:{3}, Move:{4}, SoftCopy:{5}] RFID: {6}", cleanName, savedRole, savedOffice, chkCanRoute.Checked, chkCanMove.Checked, chkCanSoftCopy.Checked, maskedUid), actionType:="USER_REGISTERED")
            End If

            lblStatusMessage.Text = String.Format("User Registered: {0} [{1} - {2}]", cleanName, savedRole, savedOffice)
            txtNewUserName.Text = ""
            txtNewUserUID.Text = ""
            PopulateStaffDropdowns()
            RefreshActiveTabGrid()
        End Sub

        ''' <summary>
        ''' Clears FailedTapCount and IsLocked for the account selected in the grid. The
        ''' per-user lock persists in tbl_Users across restarts; without this action the
        ''' documented recovery is hand-editing SQL on the server.
        ''' </summary>
        Private Sub OnUnlockSelectedUser(sender As Object, e As EventArgs)
            If CurrentUser Is Nothing Then Return
            If dgvUsers.CurrentRow Is Nothing OrElse dgvUsers.CurrentRow.DataBoundItem Is Nothing Then
                lblStatusMessage.Text = "Select the locked account in the grid first."
                Return
            End If
            Dim row = DirectCast(dgvUsers.CurrentRow.DataBoundItem, DataRowView).Row
            Dim userId As Integer = 0
            If row.Table.Columns.Contains("UserID") AndAlso Not IsDBNull(row("UserID")) Then userId = Convert.ToInt32(row("UserID"))
            Dim fullName = If(row.Table.Columns.Contains("FullName"), row("FullName").ToString(), "")
            If userId <= 0 Then
                lblStatusMessage.Text = "This account has no server record to unlock."
                Return
            End If
            If MessageBox.Show("Unlock " & fullName & " and clear its failed badge-tap counter?", "Confirm Unlock", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then Return
            If Not Program.IsDatabaseConnected OrElse AppStartup.UserRepo Is Nothing Then
                lblStatusMessage.Text = "The unlock needs the office server connection."
                Return
            End If
            Try
                AppStartup.UserRepo.ResetFailedTaps(userId)
            Catch ex As Exception
                lblStatusMessage.Text = "Unlock failed: " & ex.Message
                Return
            End Try
            EmbeddedDB.LogAudit(CurrentUser("FullName").ToString(), "Unlocked user account " & fullName & " and reset its failed-tap counter.", actionType:="USER_UNLOCKED")
            lblStatusMessage.Text = "Account unlocked: " & fullName
            If Program.Coordinator IsNot Nothing Then Program.Coordinator.RefreshFromServer("Users")
            RefreshActiveTabGrid()
        End Sub

        ''' <summary>
        ''' Revokes every active badge of the account selected in the grid. The badge stops
        ''' authenticating everywhere at the next Users pull, because the mirror projection
        ''' only carries cards whose RevokedAtUTC is still NULL.
        ''' </summary>
        Private Sub OnRevokeSelectedCard(sender As Object, e As EventArgs)
            If CurrentUser Is Nothing Then Return
            If dgvUsers.CurrentRow Is Nothing OrElse dgvUsers.CurrentRow.DataBoundItem Is Nothing Then
                lblStatusMessage.Text = "Select the account whose badge is lost in the grid first."
                Return
            End If
            Dim row = DirectCast(dgvUsers.CurrentRow.DataBoundItem, DataRowView).Row
            Dim userId As Integer = 0
            If row.Table.Columns.Contains("UserID") AndAlso Not IsDBNull(row("UserID")) Then userId = Convert.ToInt32(row("UserID"))
            Dim fullName = If(row.Table.Columns.Contains("FullName"), row("FullName").ToString(), "")
            If userId <= 0 Then
                lblStatusMessage.Text = "This account has no server record; its badge only lives on this workstation's offline store."
                Return
            End If
            If MessageBox.Show("Revoke every active badge of " & fullName & "? The badge stops working on all workstations at the next sync.", "Confirm Badge Revocation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) <> DialogResult.Yes Then Return
            If Not Program.IsDatabaseConnected OrElse AppStartup.CardService Is Nothing Then
                lblStatusMessage.Text = "Badge revocation needs the office server connection."
                Return
            End If
            Try
                Dim revoked = AppStartup.CardService.RevokeActiveCardsForUser(userId, If(CurrentUser.Table.Columns.Contains("UserID") AndAlso Not IsDBNull(CurrentUser("UserID")), Convert.ToInt32(CurrentUser("UserID")), 0), "Reported lost: revoked from User & RFID Admin")
                If revoked = 0 Then
                    lblStatusMessage.Text = "No active badge found for " & fullName & "."
                    Return
                End If
                EmbeddedDB.LogAudit(CurrentUser("FullName").ToString(), "Revoked " & revoked.ToString() & " active badge(s) of " & fullName & ".", actionType:="RFID_REVOKED")
                lblStatusMessage.Text = "Badge revoked: " & fullName & " (" & revoked.ToString() & " card(s))."
            Catch ex As Exception
                lblStatusMessage.Text = "Revocation failed: " & ex.Message
                Return
            End Try
            If Program.Coordinator IsNot Nothing Then Program.Coordinator.RefreshFromServer("Users")
            RefreshActiveTabGrid()
        End Sub
    End Class
End Namespace
