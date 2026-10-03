Option Explicit On
Option Strict On

Imports System
Imports System.Data
Imports System.Drawing
Imports System.Linq
Imports System.Text
Imports System.Windows.Forms

Namespace BTA_OSG
    Public Class FormLogin
        Inherits Form

        <System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)>
        Public Property ScannedUID As String = ""

        ' The audit trail must not claim a physical tap for a badge chosen from the list.
        <System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)>
        Public Property PickedOnScreen As Boolean = False

        Private lblScanStatus As Label
        Private cmbBadge As ComboBox
        Private btnSignIn As Button
        Private btnCancel As Button
        Private ReadOnly _wedgeBuffer As New StringBuilder()
        Private _awaitingSetupBadge As Boolean = False
        Private _idleStatusText As String = ""

        Public Sub New()
            InitializeForm()
        End Sub

        Private Sub InitializeForm()
            AppAssets.ApplyFormIcon(Me)
            Me.Text = "OSG RFID Smart Card Authentication Tap Scanner"
            Me.Size = New Size(500, 372)
            Me.MinimumSize = New Size(480, 352)
            Me.StartPosition = FormStartPosition.CenterParent
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.BackColor = CivicCalmTheme.ColorCanvas
            Me.Font = CivicCalmTheme.FontBody
            ' The reader is a keyboard wedge with no text box to land in any more, so the form
            ' itself collects the keystrokes.
            Me.KeyPreview = True
            AddHandler Me.KeyPress, AddressOf OnFormKeyPress

            Dim pnlOuter As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .RowCount = 4,
                .Padding = New Padding(20)
            }
            pnlOuter.RowStyles.Add(New RowStyle(SizeType.Absolute, 140.0F))
            pnlOuter.RowStyles.Add(New RowStyle(SizeType.Absolute, 34.0F))
            pnlOuter.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
            pnlOuter.RowStyles.Add(New RowStyle(SizeType.Absolute, 32.0F))

            ' Top Card: Scanner instructions and status
            Dim pnlCard As New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(16),
                .Margin = New Padding(0, 0, 0, 12)
            }

            Dim tblCardContent As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 2,
                .RowCount = 1,
                .BackColor = CivicCalmTheme.ColorSurface
            }
            tblCardContent.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 74.0F))
            tblCardContent.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            tblCardContent.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

            Dim picLoginLogo = AppAssets.CreateLogoPictureBox(64)
            picLoginLogo.Anchor = AnchorStyles.Left
            picLoginLogo.Margin = New Padding(0, 0, 10, 0)

            Dim pnlCardText As New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = CivicCalmTheme.ColorSurface
            }

            Dim lblIcon As New Label With {
                .Text = "RFID BADGE SCANNER : ACTIVE",
                .ForeColor = CivicCalmTheme.ColorPrimary,
                .Font = CivicCalmTheme.FontSectionHeader,
                .Dock = DockStyle.Top,
                .Height = 24
            }

            lblScanStatus = New Label With {
                .Text = "Ready to sign in. Tap your RFID card on the reader," & vbCrLf & "or pick your enrolled badge from the list" & vbCrLf & "and press Sign in.",
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .Font = CivicCalmTheme.FontBody,
                .Dock = DockStyle.Fill,
                .TextAlign = ContentAlignment.MiddleLeft
            }
            _idleStatusText = lblScanStatus.Text

            pnlCardText.Controls.Add(lblScanStatus)
            pnlCardText.Controls.Add(lblIcon)

            tblCardContent.Controls.Add(picLoginLogo, 0, 0)
            tblCardContent.Controls.Add(pnlCardText, 1, 0)
            pnlCard.Controls.Add(tblCardContent)

            ' The station lists the identities it actually holds and signs in the one picked.
            ' A physical reader is still first class: it is a keyboard wedge, so its keystrokes
            ' collect in the form buffer below, and a tap's trailing Enter submits the tapped
            ' card in preference to whatever is highlighted.
            Dim pnlPick As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 2,
                .RowCount = 1,
                .Margin = New Padding(0, 4, 0, 0)
            }
            pnlPick.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 92.0F))
            pnlPick.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            pnlPick.RowStyles.Add(New RowStyle(SizeType.Absolute, 28.0F))

            Dim lblPick As New Label With {
                .Text = "Sign in &as:",
                .Font = CivicCalmTheme.FontFieldLabel,
                .ForeColor = CivicCalmTheme.ColorInk,
                .Dock = DockStyle.Fill,
                .TextAlign = ContentAlignment.MiddleLeft,
                .Margin = New Padding(0, 0, 6, 0)
            }
            cmbBadge = New ComboBox With {
                .Dock = DockStyle.Fill,
                .DropDownStyle = ComboBoxStyle.DropDownList,
                .Font = CivicCalmTheme.FontBody,
                .BackColor = CivicCalmTheme.ColorSurface,
                .ForeColor = CivicCalmTheme.ColorInk,
                .FlatStyle = FlatStyle.Flat,
                .Margin = New Padding(0, 2, 0, 2),
                .TabIndex = 0
            }
            pnlPick.Controls.Add(lblPick, 0, 0)
            pnlPick.Controls.Add(cmbBadge, 1, 0)
            PopulateEnrolledBadges()

            Dim flwActions As New FlowLayoutPanel With {
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .FlowDirection = FlowDirection.RightToLeft,
                .WrapContents = False,
                .Margin = New Padding(0, 8, 0, 0)
            }
            btnSignIn = New Button With {
                .Text = "&Sign in",
                .Size = New Size(110, 32),
                .BackColor = CivicCalmTheme.ColorPrimary,
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .TabIndex = 1
            }
            btnSignIn.FlatAppearance.BorderSize = 0
            AddHandler btnSignIn.Click, Sub() OnSignIn()

            btnCancel = New Button With {
                .Text = "&Cancel",
                .Size = New Size(100, 32),
                .BackColor = CivicCalmTheme.ColorWell,
                .ForeColor = CivicCalmTheme.ColorInk,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .TabIndex = 2
            }
            btnCancel.FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder
            AddHandler btnCancel.Click, Sub()
                                            Me.DialogResult = DialogResult.Cancel
                                            Me.Close()
                                        End Sub
            ' RightToLeft flow places the first control at the right edge, so the primary
            ' action sits furthest right and Cancel never takes an accidental click.
            flwActions.Controls.AddRange(New Control() {btnSignIn, btnCancel})

            ' DESIGN 5.3: Enter is the only submit key the reader sends, so it belongs to Sign
            ' in. Focus starts there because a focused button handles Enter itself, which would
            ' otherwise let a dropdown steal the tap.
            Me.AcceptButton = btnSignIn
            Me.CancelButton = btnCancel

            pnlOuter.Controls.Add(pnlCard, 0, 0)
            pnlOuter.Controls.Add(pnlPick, 0, 1)
            pnlOuter.Controls.Add(flwActions, 0, 2)

            Dim lnkSetup As New LinkLabel With {
                .Text = If(AppSettings.Instance.IsConfigured AndAlso Not AppSettings.Instance.DatabaseSettings.UseSqlServer,
                           "Connect to the Office Server (Station Setup)",
                           "Station Setup & Connection Settings"),
                .Font = CivicCalmTheme.FontMicrocopy,
                .LinkColor = CivicCalmTheme.ColorPrimary,
                .ActiveLinkColor = CivicCalmTheme.ColorInfo,
                .VisitedLinkColor = CivicCalmTheme.ColorPrimary,
                .Dock = DockStyle.Fill,
                .TextAlign = ContentAlignment.MiddleCenter,
                .Cursor = Cursors.Hand,
                .TabStop = True
            }
            AddHandler lnkSetup.LinkClicked, Sub()
                ' A configured workstation must not let a logged-out user rewrite the SQL
                ' connection, reader kind, or portal switch; first-run stays open. The
                ' authorizing badge is answered right here rather than in a second dialog, so
                ' one card never feeds two stacked modals at once.
                If Not AppSettings.Instance.IsConfigured Then
                    OpenStationSetup()
                    Return
                End If
                _awaitingSetupBadge = True
                _wedgeBuffer.Clear()
                btnSignIn.Focus()
                lblScanStatus.Text = "Tap your System Administrator badge, or pick it from the list, to open Station Setup."
            End Sub

            pnlOuter.Controls.Add(lnkSetup, 0, 3)

            Me.Controls.Add(pnlOuter)
            ' Focus on Sign in so the reader's trailing Enter submits the tap, and so a stray
            ' keystroke cannot open the dropdown while a card is being read.
            Me.ActiveControl = btnSignIn
        End Sub

        Protected Overrides Sub OnShown(e As EventArgs)
            MyBase.OnShown(e)
            PopulateEnrolledBadges()
        End Sub

        ''' <summary>
        ''' Enter and the Sign in button share this. A card that was just read wins over the
        ''' highlighted list entry, because the reader's keystrokes arrive seconds before the
        ''' operator gets to the button.
        ''' </summary>
        Private Sub OnSignIn()
            Dim tapped = _wedgeBuffer.ToString()
            _wedgeBuffer.Clear()
            If tapped.Trim().Length > 0 Then
                SubmitUID(tapped)
                Return
            End If

            Dim choice = TryCast(cmbBadge.SelectedItem, BadgeChoice)
            If choice Is Nothing Then
                lblScanStatus.Text = "Pick a badge from the list, or tap a card on the reader."
                Return
            End If
            SubmitUID(choice.Uid, pickedOnScreen:=True)
        End Sub

        Private Sub OnFormKeyPress(sender As Object, e As KeyPressEventArgs)
            ' Enter belongs to Sign in; everything the reader sends is card data.
            If Not Char.IsControl(e.KeyChar) Then _wedgeBuffer.Append(e.KeyChar)
        End Sub

        ''' <summary>
        ''' Lists who may sign in on this station, taken from the store it authenticates against:
        ''' the mirrored SQL identities in office mode, the local cache otherwise. Both arrive in
        ''' the same cache table, so there is one list to build and nothing to keep in step.
        ''' </summary>
        Private Sub PopulateEnrolledBadges()
            Dim previous = If(TryCast(cmbBadge.SelectedItem, BadgeChoice), Nothing)?.Uid
            cmbBadge.BeginUpdate()
            Try
                cmbBadge.Items.Clear()
                If EmbeddedDB.DataSet.Tables.Contains("Users") Then
                    Dim dt = EmbeddedDB.DataSet.Tables("Users")
                    For Each row As DataRow In dt.Select("", "FullName ASC")
                        Dim uid = If(row.Table.Columns.Contains("RFID_UID"), row("RFID_UID").ToString(), "")
                        If uid.Length = 0 OrElse IsDBNull(row("FullName")) Then Continue For
                        Dim role = If(row.Table.Columns.Contains("Role"), row("Role").ToString(), "")
                        Dim caption = row("FullName").ToString()
                        If role.Length > 0 Then caption &= " (" & role & ")"
                        cmbBadge.Items.Add(New BadgeChoice With {.Uid = uid, .Caption = caption})
                    Next
                End If

                If previous IsNot Nothing Then
                    Dim keep = cmbBadge.Items.Cast(Of BadgeChoice)().FirstOrDefault(Function(c) c.Uid = previous)
                    If keep IsNot Nothing Then
                        cmbBadge.SelectedItem = keep
                    Else
                        cmbBadge.SelectedIndex = -1
                    End If
                ElseIf cmbBadge.Items.Count > 0 Then
                    ' A station with exactly one identity should need one keypress, not two.
                    cmbBadge.SelectedIndex = 0
                End If
            Finally
                cmbBadge.EndUpdate()
            End Try
        End Sub

        Private Sub SubmitUID(uid As String, Optional pickedOnScreen As Boolean = False)
            ' Readers pad the card ID with control characters; they must never reach the lookup.
            Dim cleanUid = New String(uid.Where(Function(c) Not Char.IsControl(c)).ToArray()).Trim().ToUpperInvariant()
            If cleanUid.Length = 0 Then Return
            PickedOnScreen = pickedOnScreen
            _wedgeBuffer.Clear()

            If _awaitingSetupBadge Then
                _awaitingSetupBadge = False
                Dim gateUser = EmbeddedDB.AuthenticateRFID(cleanUid)
                If gateUser Is Nothing Then
                    MessageBox.Show("Unrecognized card. Station setup requires a System Administrator badge.", "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                ElseIf gateUser("Role").ToString() = "System Administrator" OrElse gateUser("Role").ToString() = "SYSADMIN" Then
                    EmbeddedDB.LogAudit(gateUser("FullName").ToString(), "Authorized station setup change", actionType:="SETUP_AUTHORIZED")
                    OpenStationSetup()
                Else
                    EmbeddedDB.LogAudit(gateUser("FullName").ToString(), "Denied station setup change (role: " & gateUser("Role").ToString() & ")", actionType:="SETUP_DENIED")
                    MessageBox.Show("Only a System Administrator may change station configuration.", "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If
                lblScanStatus.Text = _idleStatusText
                Return
            End If
            ScannedUID = cleanUid
            Me.DialogResult = DialogResult.OK
            Me.Close()
        End Sub

        Private Sub OpenStationSetup()
            Using connect As New FormConnect()
                If connect.ShowDialog(Me) = DialogResult.OK Then
                    Program.ApplyStationSetup()
                    lblScanStatus.Text = "Settings updated. Ready to scan physical RFID card."
                    Return
                End If
            End Using
            lblScanStatus.Text = _idleStatusText
        End Sub

        Private Class BadgeChoice
            Public Property Uid As String
            Public Property Caption As String
            Public Overrides Function ToString() As String
                Return Caption
            End Function
        End Class
    End Class
End Namespace
