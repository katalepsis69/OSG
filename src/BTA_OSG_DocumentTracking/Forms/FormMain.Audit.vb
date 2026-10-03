Option Explicit On
Option Strict On

Imports System
Imports System.Data
Imports System.Drawing
Imports System.Threading.Tasks
Imports System.Windows.Forms

Namespace BTA_OSG
    Partial Public Class FormMain
        Private Sub SetupAuditView()
            viewAudit = New Panel With {.Dock = DockStyle.Fill, .BackColor = CivicCalmTheme.ColorCanvas}

            Dim tblAudit As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .RowCount = 2,
                .BackColor = CivicCalmTheme.ColorCanvas,
                .Padding = New Padding(0)
            }
            tblAudit.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            tblAudit.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            tblAudit.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

            Dim pnlAuditCard As New Panel With {
                .Dock = DockStyle.Fill,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .Margin = New Padding(0, 0, 0, 12),
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(12)
            }
            ApplyCardBorder(pnlAuditCard)

            Dim flwAudit As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .FlowDirection = FlowDirection.LeftToRight,
                .WrapContents = True,
                .BackColor = CivicCalmTheme.ColorSurface
            }

            Dim lblF As New Label With {
                .Text = "Filter Audit Trail:",
                .UseMnemonic = False,
                .Font = CivicCalmTheme.FontFieldLabel,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .AutoSize = True,
                .Margin = New Padding(0, 8, 8, 0)
            }

            txtAuditSearch = New TextBox With {
                .Width = CInt(Dpi(240.0F)),
                .Height = 32,
                .BackColor = CivicCalmTheme.ColorSurface,
                .ForeColor = CivicCalmTheme.ColorInk,
                .BorderStyle = BorderStyle.FixedSingle,
                .Margin = New Padding(0, 4, 8, 0)
            }
            AddHandler txtAuditSearch.KeyDown, Sub(s As Object, e As KeyEventArgs)
                                                   If e.KeyCode = Keys.Enter Then
                                                       e.SuppressKeyPress = True
                                                       OnFilterAudit(s, EventArgs.Empty)
                                                   End If
                                               End Sub

            btnAuditFilter = New Button With {
                .Text = " &Filter",
                .Size = New Size(100, 32),
                .BackColor = CivicCalmTheme.ColorPrimary,
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(0, 4, 8, 0),
                .Image = AppAssets.GetIcon("magnifying-glass", 16, Color.White),
                .ImageAlign = ContentAlignment.MiddleLeft,
                .TextAlign = ContentAlignment.MiddleCenter,
                .TextImageRelation = TextImageRelation.ImageBeforeText,
                .Padding = New Padding(6, 0, 6, 0)
            }
            btnAuditFilter.FlatAppearance.BorderSize = 0

            btnAuditReset = New Button With {
                .Text = " &Reset",
                .Size = New Size(100, 32),
                .BackColor = CivicCalmTheme.ColorWell,
                .ForeColor = CivicCalmTheme.ColorInk,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontBody,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(0, 4, 8, 0),
                .Image = AppAssets.GetIcon("arrow-clockwise", 14, CivicCalmTheme.ColorInk),
                .ImageAlign = ContentAlignment.MiddleLeft,
                .TextAlign = ContentAlignment.MiddleCenter,
                .TextImageRelation = TextImageRelation.ImageBeforeText,
                .Padding = New Padding(6, 0, 6, 0)
            }
            btnAuditReset.FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder

            ' Inline tamper-evidence: the same recomputation /verify-audit performs.
            btnAuditVerify = New Button With {
                .Text = " &Verify Chain",
                .Size = New Size(130, 32),
                .BackColor = CivicCalmTheme.ColorWell,
                .ForeColor = CivicCalmTheme.ColorInk,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontBody,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(0, 4, 0, 0),
                .Image = AppAssets.GetIcon("shield-check", 16, CivicCalmTheme.ColorInk),
                .ImageAlign = ContentAlignment.MiddleLeft,
                .TextAlign = ContentAlignment.MiddleCenter,
                .TextImageRelation = TextImageRelation.ImageBeforeText,
                .Padding = New Padding(6, 0, 6, 0)
            }
            btnAuditVerify.FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder

            AddHandler btnAuditFilter.Click, AddressOf OnFilterAudit
            AddHandler btnAuditVerify.Click, AddressOf OnVerifyAuditChain
            AddHandler btnAuditReset.Click, Sub()
                                                txtAuditSearch.Text = ""
                                                OnFilterAudit(Nothing, EventArgs.Empty)
                                            End Sub

            flwAudit.Controls.AddRange(New Control() {lblF, txtAuditSearch, btnAuditFilter, btnAuditReset, btnAuditVerify})
            pnlAuditCard.Controls.Add(flwAudit)

            Dim pnlGridCard As New Panel With {
                .Dock = DockStyle.Fill,
                .Margin = New Padding(0),
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(12)
            }
            ApplyCardBorder(pnlGridCard)

            dgvAudit = New DataGridView With {
                .Dock = DockStyle.Fill,
                .ReadOnly = True,
                .AllowUserToAddRows = False
            }
            DataGridStyler.ApplyCivicStyle(dgvAudit)

            lblAuditWatermark = CreateGridWatermark(pnlGridCard)

            pnlGridCard.Controls.Add(dgvAudit)

            tblAudit.Controls.Add(pnlAuditCard, 0, 0)
            tblAudit.Controls.Add(pnlGridCard, 0, 1)

            viewAudit.Controls.Add(tblAudit)
        End Sub

        ''' <summary>
        ''' Verifies the append-only hash chain against the office server and shows the
        ''' verdict inline. Runs off the UI thread because the recomputation walks every
        ''' sealed entry.
        ''' </summary>
        Private Sub OnVerifyAuditChain(sender As Object, e As EventArgs)
            If Not Program.IsDatabaseConnected OrElse AppStartup.AuditRepo Is Nothing Then
                lblStatusMessage.Text = "Chain verification reads the office server; connect to it first."
                Return
            End If
            btnAuditVerify.Enabled = False
            lblStatusMessage.Text = "Verifying the audit chain..."
            Dim uiScheduler = TaskScheduler.FromCurrentSynchronizationContext()
            Task.Run(Function()
                         Try
                             Return AppStartup.AuditRepo.VerifyChain()
                         Catch ex As Exception
                             Return New AuditChainResult With {.IsValid = False, .Reason = ex.Message}
                         End Try
                     End Function).ContinueWith(
                Sub(t)
                    If Me.IsDisposed OrElse Me.Disposing Then Return
                    btnAuditVerify.Enabled = True
                    lblStatusMessage.ForeColor = CivicCalmTheme.ColorInk
                    Dim r = t.Result
                    If r.IsValid Then
                        lblStatusMessage.Text = "Audit chain verified: " & r.RowsChecked.ToString() & " sealed entries intact" &
                            If(r.RowsUnsealed > 0, " (" & r.RowsUnsealed.ToString() & " not sealed yet).", ".")
                        lblStatusMessage.ForeColor = CivicCalmTheme.ColorPrimary
                    Else
                        lblStatusMessage.Text = "AUDIT CHAIN BROKEN: " & r.Reason
                        lblStatusMessage.ForeColor = CivicCalmTheme.ColorDanger
                    End If
                End Sub, uiScheduler)
        End Sub

        Private Sub OnFilterAudit(sender As Object, e As EventArgs)
            Dim rawQuery = txtAuditSearch.Text.Trim()
            Dim dt = EmbeddedDB.DataSet.Tables("AuditTrail")
            If String.IsNullOrEmpty(rawQuery) Then
                Dim dv As New DataView(dt) With {.Sort = "AuditID DESC"}
                dgvAudit.DataSource = dv
                DataGridStyler.FormatAuditColumns(dgvAudit)
                lblStatusMessage.Text = "Displaying complete audit trail."
            Else
                Try
                    Dim safeQ = EscapeRowFilter(rawQuery)
                    Dim dv As New DataView(dt) With {.Sort = "AuditID DESC"}
                    dv.RowFilter = String.Format("UserName LIKE '%{0}%' OR ActionDescription LIKE '%{0}%'", safeQ)
                    dgvAudit.DataSource = dv
                    DataGridStyler.FormatAuditColumns(dgvAudit)
                    lblStatusMessage.Text = String.Format("Audit log filtered for '{0}'.", rawQuery)
                Catch ex As Exception
                    dgvAudit.DataSource = dt.Clone()
                    DataGridStyler.FormatAuditColumns(dgvAudit)
                    lblStatusMessage.Text = "Invalid audit filter syntax entered."
                End Try
            End If

            If dgvAudit.Rows.Count = 0 Then
                DataGridStyler.SetEmptyState(dgvAudit, lblAuditWatermark, "No audit trail records match the search filter.")
            Else
                DataGridStyler.SetPopulatedState(dgvAudit, lblAuditWatermark)
            End If
            lblStatusCount.Text = dgvAudit.Rows.Count.ToString() & " Records"
        End Sub
    End Class
End Namespace
