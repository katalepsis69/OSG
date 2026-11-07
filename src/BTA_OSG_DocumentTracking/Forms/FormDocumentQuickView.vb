Option Explicit On
Option Strict On

Imports System
Imports System.Data
Imports System.Drawing
Imports System.Windows.Forms

Namespace BTA_OSG
    ''' <summary>
    ''' Compact read-only peek at a document: identity, status, custody, deadline, and
    ''' the last action, without paying the full 1180x820 detail dialog for a glance.
    ''' Enter (or the Full Details button) hands back DialogResult.OK so the caller can
    ''' open the complete detail dialog; Esc closes.
    ''' </summary>
    Public Class FormDocumentQuickView
        Inherits Form

        Public Sub New(docId As Integer)
            AppAssets.ApplyFormIcon(Me)
            Me.Text = "Quick View"
            Me.Size = New Size(520, 460)
            Me.MinimumSize = New Size(480, 420)
            Me.StartPosition = FormStartPosition.CenterParent
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.BackColor = CivicCalmTheme.ColorCanvas
            Me.Font = CivicCalmTheme.FontBody

            Dim rows As DataRow()
            ' The replay worker mutates these tables on the sync thread; read under the
            ' store lock like every other cross-thread select (same shape as the detail dialog).
            SyncLock EmbeddedDB.SyncRoot
                rows = EmbeddedDB.DataSet.Tables("Documents").Select("DocumentID = " & docId)
            End SyncLock
            If rows.Length = 0 Then
                MessageBox.Show("Document record not found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.Close()
                Return
            End If
            Dim doc As DataRow = rows(0)

            Dim pnlCard As New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(20),
                .Margin = New Padding(14)
            }

            Dim lblCode As New Label With {
                .Text = doc("DocCode").ToString(),
                .Font = CivicCalmTheme.FontIdentifier,
                .ForeColor = CivicCalmTheme.ColorPrimary,
                .Dock = DockStyle.Top,
                .Height = 30,
                .TextAlign = ContentAlignment.MiddleLeft
            }

            Dim lblTitle As New Label With {
                .Text = doc("Title").ToString(),
                .Font = CivicCalmTheme.FontFormTitle,
                .ForeColor = CivicCalmTheme.ColorInk,
                .Dock = DockStyle.Top,
                .Height = 56,
                .TextAlign = ContentAlignment.MiddleLeft
            }

            Dim status As String = doc("CurrentStatus").ToString()
            Dim lblStatus As New Label With {
                .Text = "Status: " & DocumentStatus.DisplayName(status),
                .Font = CivicCalmTheme.FontFieldLabel,
                .Dock = DockStyle.Top,
                .Height = 30,
                .TextAlign = ContentAlignment.MiddleLeft,
                .Padding = New Padding(6, 4, 0, 0)
            }
            ApplyStatusTint(lblStatus, status)

            Dim dueText = doc("TargetDeadlineUTC").ToString().Trim()
            Dim dueValue As DateTime
            Dim dueColor = CivicCalmTheme.ColorInk
            Dim dueFont = CivicCalmTheme.FontBody
            If DateTime.TryParse(dueText, Globalization.CultureInfo.InvariantCulture, Globalization.DateTimeStyles.None, dueValue) AndAlso
               Not DocumentStatus.IsTerminalStatus(status) Then
                If dueValue < DateTime.Now Then
                    dueText &= "  (overdue)"
                    dueColor = CivicCalmTheme.ColorDanger
                    dueFont = CivicCalmTheme.FontFieldLabel
                ElseIf dueValue <= DateTime.Now.AddHours(24) Then
                    dueText &= "  (due within 24 h)"
                    dueColor = CivicCalmTheme.ColorStatusReceivedFg
                    dueFont = CivicCalmTheme.FontFieldLabel
                End If
            End If

            Dim tbl As New TableLayoutPanel With {
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .ColumnCount = 2,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(0, 8, 0, 0)
            }
            tbl.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 130.0F))
            tbl.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            AddDetailRow(tbl, 0, "Assigned Desk:", doc("AssignedSection").ToString())
            AddDetailRow(tbl, 1, "Assigned Staff:", doc("AssignedStaff").ToString())
            AddDetailRow(tbl, 2, "Origin Office:", doc("OriginatingOffice").ToString())
            AddDetailRow(tbl, 3, "Destination:", doc("DestinationOffice").ToString())
            AddDetailRow(tbl, 4, "Target Deadline:", dueText, dueColor, dueFont)
            AddDetailRow(tbl, 5, "Storage:", String.Format("{0} / {1} / {2}", doc("CabinetID"), doc("ShelfNo"), doc("BoxCode")))

            Dim lblLastAction As New Label With {
                .Text = "Last action: " & doc("LastActionTaken").ToString(),
                .Font = CivicCalmTheme.FontMicrocopy,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .Dock = DockStyle.Top,
                .Height = 44,
                .TextAlign = ContentAlignment.MiddleLeft
            }

            Dim flwButtons As New FlowLayoutPanel With {
                .Dock = DockStyle.Bottom,
                .FlowDirection = FlowDirection.RightToLeft,
                .WrapContents = False,
                .Height = 52,
                .BackColor = CivicCalmTheme.ColorSurface
            }
            Dim btnClose = CivicCalmTheme.CreateIconButton("&Close", "x", 14, minHeight:=32, margin:=New Padding(8, 8, 0, 8))
            AddHandler btnClose.Click, Sub()
                                           Me.DialogResult = DialogResult.Cancel
                                           Me.Close()
                                       End Sub
            Dim btnFull = CivicCalmTheme.CreateIconButton("&Full Details", "file-text", 14, backColor:=CivicCalmTheme.ColorPrimary, minHeight:=32)
            AddHandler btnFull.Click, Sub()
                                          Me.DialogResult = DialogResult.OK
                                          Me.Close()
                                      End Sub
            flwButtons.Controls.AddRange(New Control() {btnClose, btnFull})

            pnlCard.Controls.Add(lblLastAction)
            pnlCard.Controls.Add(tbl)
            pnlCard.Controls.Add(lblStatus)
            pnlCard.Controls.Add(lblTitle)
            pnlCard.Controls.Add(lblCode)
            Me.Controls.Add(pnlCard)
            Me.Controls.Add(flwButtons)

            Me.AcceptButton = btnFull
            Me.CancelButton = btnClose
        End Sub

        Private Sub AddDetailRow(tbl As TableLayoutPanel, rowIndex As Integer, caption As String, value As String,
                                 Optional valueColor As Color = Nothing, Optional valueFont As Font = Nothing)
            Dim lblCaption As New Label With {
                .Text = caption,
                .Font = CivicCalmTheme.FontFieldLabel,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .Dock = DockStyle.Fill,
                .TextAlign = ContentAlignment.MiddleLeft,
                .Margin = New Padding(0, 2, 8, 2)
            }
            Dim lblValue As New Label With {
                .Text = value,
                .Font = If(valueFont, CivicCalmTheme.FontBody),
                .ForeColor = If(valueColor = Color.Empty, CivicCalmTheme.ColorInk, valueColor),
                .Dock = DockStyle.Fill,
                .TextAlign = ContentAlignment.MiddleLeft,
                .AutoSize = True,
                .Margin = New Padding(0, 2, 0, 2)
            }
            tbl.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            tbl.Controls.Add(lblCaption, 0, rowIndex)
            tbl.Controls.Add(lblValue, 1, rowIndex)
        End Sub

        Private Sub ApplyStatusTint(lbl As Label, status As String)
            ' One status-color vocabulary: the pairs come from DataGridStyler so the
            ' quick view and every grid tint the same code identically.
            Dim backColor As Color = Nothing
            Dim foreColor As Color = Nothing
            If DataGridStyler.StatusColors(status, backColor, foreColor) Then
                lbl.BackColor = backColor
                lbl.ForeColor = foreColor
            Else
                lbl.BackColor = CivicCalmTheme.ColorStatusPendingBg
                lbl.ForeColor = CivicCalmTheme.ColorStatusPendingFg
            End If
        End Sub
    End Class
End Namespace
