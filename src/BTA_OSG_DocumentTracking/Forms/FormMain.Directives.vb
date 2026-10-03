Option Explicit On
Option Strict On

Imports System
Imports System.Drawing
Imports System.Windows.Forms

Namespace BTA_OSG
    Partial Public Class FormMain
        Private Sub SetupDirectivesView()
            viewDirectives = New Panel With {.Dock = DockStyle.Fill, .BackColor = CivicCalmTheme.ColorCanvas}

            Dim tblDirectives As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .RowCount = 2,
                .BackColor = CivicCalmTheme.ColorCanvas,
                .Padding = New Padding(0)
            }
            tblDirectives.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            tblDirectives.RowStyles.Add(New RowStyle(SizeType.Absolute, 116.0F))
            tblDirectives.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

            Dim pnlTopCard As New Panel With {
                .Dock = DockStyle.Fill,
                .Margin = New Padding(0, 0, 0, 12),
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(16)
            }
            ApplyCardBorder(pnlTopCard)

            Dim tblDir As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 5,
                .RowCount = 2
            }
            tblDir.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 25.0F))
            tblDir.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 22.0F))
            tblDir.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 18.0F))
            tblDir.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 20.0F))
            tblDir.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 15.0F))

            cmbDirDocs = New ComboBox With {.Dock = DockStyle.Fill, .DropDownStyle = ComboBoxStyle.DropDownList, .DropDownWidth = 420, .BackColor = CivicCalmTheme.ColorSurface, .ForeColor = CivicCalmTheme.ColorInk}
            cmbDirective = New ComboBox With {.Dock = DockStyle.Fill, .DropDownStyle = ComboBoxStyle.DropDownList, .BackColor = CivicCalmTheme.ColorSurface, .ForeColor = CivicCalmTheme.ColorInk}
            cmbDirective.Items.AddRange(New Object() {"For Immediate Action", "Referred to Committee on Rules", "Forwarded for Speaker Signature", "Under OSG Administrative Review", "Approved & Archived"})
            cmbDirective.SelectedIndex = 0

            cmbDirAssign = New ComboBox With {.Dock = DockStyle.Fill, .DropDownStyle = ComboBoxStyle.DropDownList, .BackColor = CivicCalmTheme.ColorSurface, .ForeColor = CivicCalmTheme.ColorInk}
            txtDirNotes = New TextBox With {.Dock = DockStyle.Fill, .Text = "", .BackColor = CivicCalmTheme.ColorSurface, .ForeColor = CivicCalmTheme.ColorInk, .BorderStyle = BorderStyle.FixedSingle}

            btnApplyDirective = New Button With {
                .Text = " &Log Action Directive",
                .Dock = DockStyle.Fill,
                .BackColor = CivicCalmTheme.ColorPrimary,
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .Image = AppAssets.GetIcon("sparkle", 16, Color.White),
                .ImageAlign = ContentAlignment.MiddleLeft,
                .TextAlign = ContentAlignment.MiddleCenter,
                .TextImageRelation = TextImageRelation.ImageBeforeText,
                .Padding = New Padding(6, 0, 6, 0)
            }
            btnApplyDirective.FlatAppearance.BorderSize = 0
            AddHandler btnApplyDirective.Click, AddressOf OnApplyDirective

            tblDir.Controls.Add(New Label With {.Text = "Select Document:", .UseMnemonic = False, .Font = CivicCalmTheme.FontFieldLabel, .ForeColor = CivicCalmTheme.ColorInkMuted, .Dock = DockStyle.Fill}, 0, 0)
            tblDir.Controls.Add(New Label With {.Text = "SG Directive:", .UseMnemonic = False, .Font = CivicCalmTheme.FontFieldLabel, .ForeColor = CivicCalmTheme.ColorInkMuted, .Dock = DockStyle.Fill}, 1, 0)
            tblDir.Controls.Add(New Label With {.Text = "Reassign Staff:", .UseMnemonic = False, .Font = CivicCalmTheme.FontFieldLabel, .ForeColor = CivicCalmTheme.ColorInkMuted, .Dock = DockStyle.Fill}, 2, 0)
            tblDir.Controls.Add(New Label With {.Text = "Directive Notes / Remarks:", .UseMnemonic = False, .Font = CivicCalmTheme.FontFieldLabel, .ForeColor = CivicCalmTheme.ColorInkMuted, .Dock = DockStyle.Fill}, 3, 0)
            tblDir.Controls.Add(New Label With {.Text = "", .Dock = DockStyle.Fill}, 4, 0)

            tblDir.Controls.Add(cmbDirDocs, 0, 1)
            tblDir.Controls.Add(cmbDirective, 1, 1)
            tblDir.Controls.Add(cmbDirAssign, 2, 1)
            tblDir.Controls.Add(txtDirNotes, 3, 1)
            tblDir.Controls.Add(btnApplyDirective, 4, 1)

            pnlTopCard.Controls.Add(tblDir)

            Dim pnlGridCard As New Panel With {
                .Dock = DockStyle.Fill,
                .Margin = New Padding(0),
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(12)
            }
            ApplyCardBorder(pnlGridCard)

            dgvDirectives = New DataGridView With {
                .Dock = DockStyle.Fill,
                .ReadOnly = True,
                .AllowUserToAddRows = False
            }
            DataGridStyler.ApplyCivicStyle(dgvDirectives)

            lblDirectivesWatermark = CreateGridWatermark(pnlGridCard)

            pnlGridCard.Controls.Add(dgvDirectives)

            tblDirectives.Controls.Add(pnlTopCard, 0, 0)
            tblDirectives.Controls.Add(pnlGridCard, 0, 1)

            viewDirectives.Controls.Add(tblDirectives)
        End Sub

        Private Sub OnApplyDirective(sender As Object, e As EventArgs)
            If CurrentUser Is Nothing Then
                lblStatusMessage.Text = "Authentication Required to Issue Directives."
                MessageBox.Show("RFID Authentication Required.", "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' Directives are an executive action (some rewrite document state), so they are
            ' gated like approve/release: SG, System Administrator, or OSG Chief.
            Dim role = CurrentUser("Role").ToString()
            If role <> "Secretary-General" AndAlso role <> "System Administrator" AndAlso role <> "SYSADMIN" AndAlso role <> "OSG Chief" AndAlso role <> "OSG_CHIEF" Then
                lblStatusMessage.Text = "Access Denied: Only the Secretary-General may issue action directives."
                MessageBox.Show("Only the Secretary-General (or a System Administrator) may log action directives.", "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            If cmbDirDocs.SelectedItem Is Nothing Then
                lblStatusMessage.Text = "Validation Error: Please select a document."
                Return
            End If

            ' A queued second Click re-validates successfully because the selections persist,
            ' so the button is disabled until the write (and its UI refresh) has completed.
            If Not btnApplyDirective.Enabled Then Return
            btnApplyDirective.Enabled = False
            Try
                Dim docStr = cmbDirDocs.SelectedItem.ToString()
                Dim docId As Integer = 0
                Dim colonIdx = docStr.IndexOf(":"c)
                If colonIdx > 0 Then
                    Dim idPart = docStr.Substring(0, colonIdx).Replace("ID ", "").Trim()
                    Integer.TryParse(idPart, docId)
                End If
                If docId <= 0 Then
                    lblStatusMessage.Text = "Validation Error: Please select a valid document."
                    Return
                End If

                Dim directive = cmbDirective.SelectedItem.ToString()
                Dim assign = If(cmbDirAssign.SelectedItem IsNot Nothing, cmbDirAssign.SelectedItem.ToString(), "")

                Dim currentUserId As Integer = 1
                If CurrentUser IsNot Nothing AndAlso CurrentUser.Table.Columns.Contains("UserID") AndAlso Not IsDBNull(CurrentUser("UserID")) Then
                    currentUserId = Convert.ToInt32(CurrentUser("UserID"))
                End If
                Dim currentUserName As String = If(CurrentUser IsNot Nothing, CurrentUser("FullName").ToString(), "System Staff")

                If Program.Coordinator IsNot Nothing Then
                    Program.Coordinator.ApplyDirective(docId, directive, assign, txtDirNotes.Text.Trim(), currentUserName, currentUserId)
                Else
                    EmbeddedDB.AddDirective(docId, directive, assign, txtDirNotes.Text.Trim(), currentUserName, directiveCode:=DesktopDataCoordinator.DirectiveCodeFor(directive))
                    EmbeddedDB.LogAudit(currentUserName, String.Format("Applied SG Directive [{0}] to Doc ID #{1}", directive, docId), actionType:="DIRECTIVE_ADDED")
                End If

                lblStatusMessage.Text = String.Format("Action Directive Logged: [{0}] applied to Doc ID #{1}", directive, docId)
                txtDirNotes.Text = ""
                RefreshActiveTabGrid()
            Finally
                btnApplyDirective.Enabled = True
            End Try
        End Sub
    End Class
End Namespace
