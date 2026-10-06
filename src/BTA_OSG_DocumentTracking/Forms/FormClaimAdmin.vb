Option Explicit On
Option Strict On

Imports System
Imports System.Drawing
Imports System.Linq
Imports System.Windows.Forms

Namespace BTA_OSG
    ''' <summary>
    ''' The one identity step of a fresh office install. A provisioned server used to carry a
    ''' published bootstrap card, which meant the first badge of a known value opened any
    ''' database in the field. This form asks who the administrator is and enrols that card, so
    ''' nobody can sign in until a real person has been named on this machine.
    ''' </summary>
    Public Class FormClaimAdmin
        Inherits Form

        Private lblStatus As Label
        Private txtFullName As TextBox
        Private txtCardUid As TextBox
        Private txtDesk As TextBox
        Private btnClaim As Button
        Private btnCancel As Button
        Private epValidation As ErrorProvider
        Private _busy As Boolean = False

        Public Sub New()
            InitializeForm()
        End Sub

        Private Sub InitializeForm()
            AppAssets.ApplyFormIcon(Me)
            Me.Text = "BTA OSG Document Tracking : System Administrator"
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.BackColor = CivicCalmTheme.ColorCanvas
            Me.Font = CivicCalmTheme.FontBody
            Me.ClientSize = New Size(560, 372)

            Dim pnlRoot As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .Padding = New Padding(20)
            }
            For i As Integer = 0 To 4
                pnlRoot.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            Next
            pnlRoot.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

            Dim pnlHeader As New TableLayoutPanel With {
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .ColumnCount = 2
            }
            pnlHeader.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 58.0F))
            pnlHeader.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            Dim picLogo = AppAssets.CreateLogoPictureBox(48)
            picLogo.Anchor = AnchorStyles.Left
            picLogo.Margin = New Padding(0, 0, 10, 0)
            Dim pnlHeaderText As New Panel With {
                .Dock = DockStyle.Fill,
                .AutoSize = True,
                .Padding = New Padding(0, 4, 0, 0)
            }
            Dim lblTitle As New Label With {
                .Text = "Claim the first administrator",
                .Font = CivicCalmTheme.FontFormTitle,
                .ForeColor = CivicCalmTheme.ColorPrimary,
                .Dock = DockStyle.Top,
                .AutoSize = True
            }
            Dim lblSubtitle As New Label With {
                .Text = "This station has no System Administrator yet.",
                .Font = CivicCalmTheme.FontFieldLabel,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .Dock = DockStyle.Top,
                .AutoSize = True
            }
            pnlHeaderText.Controls.AddRange(New Control() {lblSubtitle, lblTitle})
            pnlHeader.Controls.Add(picLogo, 0, 0)
            pnlHeader.Controls.Add(pnlHeaderText, 1, 0)

            lblStatus = New Label With {
                .Text = "Name the officer who will administer this station, then tap the badge that officer will sign in with. That card is the only credential this station will accept.",
                .Font = CivicCalmTheme.FontBody,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .AutoSize = False,
                .Dock = DockStyle.Fill,
                .Height = 62,
                .Margin = New Padding(0, 10, 0, 0)
            }

            Dim pnlFields As New TableLayoutPanel With {
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .ColumnCount = 2,
                .Padding = New Padding(0, 8, 0, 0)
            }
            pnlFields.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 34.0F))
            pnlFields.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 66.0F))

            txtFullName = NewInput(0)
            txtCardUid = NewInput(1)
            txtDesk = NewInput(2)
            txtDesk.Text = "ICT / Systems Administration"

            AddFieldRow(pnlFields, 0, "&Full name:", txtFullName)
            AddFieldRow(pnlFields, 1, "RFID &card UID:", txtCardUid)
            AddFieldRow(pnlFields, 2, "&Desk or department:", txtDesk)

            Dim pnlHint As New Label With {
                .Text = "Hold the badge to the reader, or type its card ID and press Enter.",
                .Font = CivicCalmTheme.FontMicrocopy,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .Dock = DockStyle.Fill,
                .AutoSize = True,
                .Margin = New Padding(0, 6, 0, 0)
            }

            pnlRoot.Controls.Add(pnlHeader, 0, 0)
            pnlRoot.Controls.Add(lblStatus, 0, 1)
            pnlRoot.Controls.Add(pnlFields, 0, 2)
            pnlRoot.Controls.Add(pnlHint, 0, 3)

            Dim pnlFooter As New FlowLayoutPanel With {
                .Dock = DockStyle.Bottom,
                .AutoSize = True,
                .FlowDirection = FlowDirection.RightToLeft,
                .WrapContents = False,
                .Padding = New Padding(0, 12, 0, 6)
            }
            btnCancel = New Button With {
                .Text = "Close &without claiming",
                .Size = New Size(180, 32),
                .BackColor = CivicCalmTheme.ColorWell,
                .ForeColor = CivicCalmTheme.ColorInk,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .TabIndex = 10
            }
            btnClaim = New Button With {
                .Text = "Cl&aim administrator",
                .Size = New Size(170, 32),
                .BackColor = CivicCalmTheme.ColorPrimary,
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .TabIndex = 9
            }
            btnClaim.FlatAppearance.BorderSize = 0
            btnClaim.FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder
            AddHandler btnClaim.Click, AddressOf OnClaim
            AddHandler btnCancel.Click, Sub()
                                           Me.DialogResult = DialogResult.Cancel
                                           Me.Close()
                                       End Sub
            pnlFooter.Controls.AddRange(New Control() {btnClaim, btnCancel})

            Me.Controls.Add(pnlRoot)
            Me.Controls.Add(pnlFooter)
            Me.AcceptButton = btnClaim
            Me.CancelButton = btnCancel
            epValidation = New ErrorProvider()
        End Sub

        Private Shared Function NewInput(tabIndex As Integer) As TextBox
            Return New TextBox With {
                .Dock = DockStyle.Fill,
                .Margin = New Padding(0, 2, 0, 2),
                .BackColor = CivicCalmTheme.ColorSurface,
                .ForeColor = CivicCalmTheme.ColorInk,
                .BorderStyle = BorderStyle.FixedSingle,
                .TabIndex = tabIndex
            }
        End Function

        Private Shared Sub AddFieldRow(tbl As TableLayoutPanel, row As Integer, caption As String, input As TextBox)
            tbl.Controls.Add(New Label With {
                .Text = caption,
                .Font = CivicCalmTheme.FontFieldLabel,
                .ForeColor = CivicCalmTheme.ColorInk,
                .Dock = DockStyle.Fill,
                .TextAlign = ContentAlignment.MiddleLeft,
                .Margin = New Padding(0, 2, 8, 2)
            }, 0, row)
            tbl.Controls.Add(input, 1, row)
        End Sub

        Protected Overrides Sub OnShown(e As EventArgs)
            MyBase.OnShown(e)
            txtFullName.Focus()
        End Sub

        Private Sub OnClaim(sender As Object, e As EventArgs)
            If _busy Then Return
            epValidation.Clear()

            ' The dialog can sit open while another seat claims; a second administrator
            ' would defeat the one-naming-step contract, so re-ask the store right here.
            If Program.AdministratorExists() Then
                SetStatus("An administrator has already been claimed (possibly at another seat). This station does not need a second claim.", CivicCalmTheme.ColorWarning)
                btnClaim.Enabled = False
                Return
            End If

            Dim fullName = txtFullName.Text.Trim()
            Dim cardUid = New String(txtCardUid.Text.Where(Function(c) Not Char.IsControl(c)).ToArray()).Trim().ToUpperInvariant()
            Dim desk = txtDesk.Text.Trim()
            If desk.Length = 0 Then desk = "ICT / Systems Administration"

            If fullName.Length = 0 Then
                epValidation.SetError(txtFullName, "Enter the name of the officer who will administer this station.")
                SetStatus("A full name is required.", CivicCalmTheme.ColorWarning)
                txtFullName.Focus()
                Return
            End If
            If cardUid.Length = 0 Then
                epValidation.SetError(txtCardUid, "Tap a badge, or type the UID printed on it.")
                SetStatus("The administrator signs in with a card, so a card UID is required.", CivicCalmTheme.ColorWarning)
                txtCardUid.Focus()
                Return
            End If

            ' Claiming a card that already belongs to somebody would rewrite that person's name
            ' and hand them the administrator role, so this step only ever enrols a spare card.
            Dim connected = Program.IsDatabaseConnected
            Dim owner As String
            Try
                owner = ExistingCardOwner(cardUid, connected)
            Catch ex As Exception
                SetStatus(If(connected, "The office server could not be read", "The local store could not be read") &
                          " (" & ex.GetType().Name & "). Try again once it answers.", CivicCalmTheme.ColorDanger)
                Return
            End Try
            If owner.Length > 0 Then
                epValidation.SetError(txtCardUid, "Already enrolled for " & owner & ".")
                SetStatus("That card is already enrolled for " & owner & ". Enrol the administrator from a signed-in Admin tab, or tap a different card.", CivicCalmTheme.ColorWarning)
                txtCardUid.Focus()
                Return
            End If

            SetBusy(True)
            Try
                Dim claimed = Program.Coordinator.RegisterUser(fullName, "System Administrator", desk, cardUid,
                                                               canRoute:=True, canMove:=True, canSoftCopy:=True,
                                                               enrolledByName:=fullName, enrolledByUserId:=0,
                                                               claimGuard:=True)
                If connected AndAlso Not claimed Then
                    ' RegisterUser keeps the claim in the local store (pendingSync) when SQL
                    ' refuses it, so "nothing was saved" would be false, and a retry, often
                    ' with a different card, would replay two administrators.
                    Dim queued = EmbeddedDB.DataSet.Tables("Users").Select("PendingSync = True").Length > 0
                    If queued Then
                        SetStatus("The office server did not accept this enrolment, but the claim is saved on this workstation and will sync automatically. Do not claim again here or at another seat.", CivicCalmTheme.ColorDanger)
                    Else
                        SetStatus("The office server did not accept this enrolment. Nothing was saved; check the connection in Station Setup and try again.", CivicCalmTheme.ColorDanger)
                    End If
                    Return
                End If
                If connected Then
                    LogClaim(fullName, cardUid)
                Else
                    ' FormMain re-reads the cache file on construction, so an unflushed claim
                    ' would vanish before the first badge tap.
                    EmbeddedDB.Save()
                End If
                Me.DialogResult = DialogResult.OK
                Me.Close()
            Catch refused As GuardRefusedException
                ' Another SYSADMIN exists: the server refused the claim inside its
                ' transaction. Surface it and cache nothing, so a retry cannot stack a
                ' second administrator into the replay outbox.
                SetStatus(refused.Message, CivicCalmTheme.ColorDanger)
                Return
            Finally
                SetBusy(False)
            End Try
        End Sub

        Private Shared Function ExistingCardOwner(cardUid As String, connected As Boolean) As String
            If connected Then
                Dim found = AppStartup.UserRepo.GetByCardPublicID(cardUid)
                Return If(found IsNot Nothing, found.FullName, "")
            End If
            Dim rows = EmbeddedDB.DataSet.Tables("Users").Select(String.Format(
                "RFID_UID_HASH = '{0}' OR RFID_UID = '{1}' OR RFID_UID_RAW = '{1}'",
                EmbeddedDB.Sha256Hex(cardUid).Replace("'", "''"), cardUid.Replace("'", "''")))
            Return If(rows.Length > 0, rows(0)("FullName").ToString(), "")
        End Function

        Private Sub LogClaim(fullName As String, cardUid As String)
            Dim masked = If(cardUid.Length > 4, "****" & cardUid.Substring(cardUid.Length - 4), cardUid)
            Try
                AppStartup.AuditService.LogEvent("ADMIN_CLAIMED", "User", Nothing, Nothing, Nothing,
                                                 "System Administrator " & fullName & " enrolled with card " & masked, True, Nothing)
            Catch ex As Exception
                ' The enrolment is the fact that matters; a lost audit row must not read as a
                ' failed claim and invite a second, conflicting attempt.
                Program.ReportOperatorWarning("Administrator claim was not audited: " & ex.Message)
            End Try
        End Sub

        Private Sub SetBusy(busy As Boolean)
            _busy = busy
            btnClaim.Enabled = Not busy
            btnCancel.Enabled = Not busy
            txtFullName.Enabled = Not busy
            txtCardUid.Enabled = Not busy
            txtDesk.Enabled = Not busy
            If busy Then SetStatus("Enrolling the administrator...", CivicCalmTheme.ColorInfo)
        End Sub

        Private Sub SetStatus(text As String, color As Color)
            lblStatus.Text = text
            lblStatus.ForeColor = color
        End Sub
    End Class
End Namespace
