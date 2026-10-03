Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Drawing.Printing
Imports System.Windows.Forms

Namespace BTA_OSG
    Public Class RoutingSlipPrintService
        Private ReadOnly _doc As Document
        Private ReadOnly _steps As List(Of RoutingStep)
        Private ReadOnly _directives As List(Of ActionDirective)

        Public Sub New(doc As Document, Optional routingLogs As List(Of RoutingLog) = Nothing, Optional directives As List(Of ActionDirective) = Nothing)
            _doc = doc
            _steps = RoutingStepService.GetStepsForDocument(doc, routingLogs)
            _directives = If(directives, New List(Of ActionDirective)())
        End Sub

        Public Shared Sub ShowPreview(doc As Document, Optional routingLogs As List(Of RoutingLog) = Nothing, Optional directives As List(Of ActionDirective) = Nothing, Optional owner As Form = Nothing)
            If doc Is Nothing Then
                MessageBox.Show("No document selected for routing slip generation.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim service As New RoutingSlipPrintService(doc, routingLogs, directives)
            Using printDoc As New PrintDocument()
                printDoc.DocumentName = "RoutingSlip_" & doc.DocCode
                ConfigurePage(printDoc)
                AddHandler printDoc.PrintPage, AddressOf service.OnPrintPage

                ' The stock PrintPreviewDialog leaves its toolbar and preview pane inert on the
                ' office workstations (no print, zoom, or scroll input gets through), so the slip
                ' previews in our own form where every control is wired here.
                Using preview As New RoutingSlipPreviewForm(printDoc)
                    preview.Text = "Official Routing Slip Preview: " & doc.DocCode
                    If owner IsNot Nothing Then
                        preview.ShowDialog(owner)
                    Else
                        preview.ShowDialog()
                    End If
                End Using
            End Using
        End Sub

        Public Shared Sub PrintDirect(doc As Document, Optional routingLogs As List(Of RoutingLog) = Nothing, Optional directives As List(Of ActionDirective) = Nothing, Optional owner As Form = Nothing)
            If doc Is Nothing Then Return

            Dim service As New RoutingSlipPrintService(doc, routingLogs, directives)
            Using printDoc As New PrintDocument()
                printDoc.DocumentName = "RoutingSlip_" & doc.DocCode
                AddHandler printDoc.PrintPage, AddressOf service.OnPrintPage

                Using printDlg As New PrintDialog()
                    printDlg.Document = printDoc
                    printDlg.UseEXDialog = True
                    Dim result = If(owner IsNot Nothing, printDlg.ShowDialog(owner), printDlg.ShowDialog())
                    If result = DialogResult.OK Then
                        ' Applied after the dialog: assigning PrinterSettings rebuilds DefaultPageSettings
                        ' from the chosen device's devmode and discards anything set beforehand.
                        ConfigurePage(printDoc)
                        printDoc.Print()
                    End If
                End Using
            End Using
        End Sub

        ' The slip is laid out for A4. Page geometry is pinned rather than inherited because a
        ' receipt or label driver reports a paper too small to hold six columns, and its driver
        ' may not report a size at all.
        Friend Shared Sub ConfigurePage(printDoc As PrintDocument)
            printDoc.DefaultPageSettings.PaperSize = New PaperSize("A4", 827, 1169)
            printDoc.DefaultPageSettings.Margins = New Margins(40, 40, 40, 40)
            printDoc.DefaultPageSettings.Landscape = False
        End Sub

        Private Sub OnPrintPage(sender As Object, e As PrintPageEventArgs)
            Dim g = e.Graphics
            g.SmoothingMode = SmoothingMode.AntiAlias
            g.TextRenderingHint = Drawing.Text.TextRenderingHint.ClearTypeGridFit

            Dim realBounds = e.MarginBounds
            Const designWidth As Single = 747.0F
            Dim fitScale As Single = Math.Min(1.0F, realBounds.Width / designWidth)
            Dim pageState = g.Save()
            g.ScaleTransform(fitScale, fitScale)
            Try
                Dim bounds As New RectangleF(realBounds.Left, realBounds.Top, designWidth, realBounds.Height / fitScale)
                Dim currentY As Single = bounds.Top

                Using fontHeaderSmall As New Font("Segoe UI", 9.0F, FontStyle.Regular),
                      fontHeaderMain As New Font("Segoe UI", 12.0F, FontStyle.Bold),
                      fontHeaderDoc As New Font("Segoe UI", 14.0F, FontStyle.Bold),
                      fontSectionTitle As New Font("Segoe UI", 9.5F, FontStyle.Bold),
                      fontBold As New Font("Segoe UI", 8.5F, FontStyle.Bold),
                      fontRegular As New Font("Segoe UI", 8.5F, FontStyle.Regular),
                      fontSmall As New Font("Segoe UI", 7.5F, FontStyle.Regular),
                      penBorder As New Pen(Color.FromArgb(40, 50, 65), 1.0F),
                      penThick As New Pen(Color.FromArgb(20, 30, 45), 2.0F),
                      penDotted As New Pen(Color.FromArgb(160, 170, 180), 1.0F) With {.DashStyle = DashStyle.Dot},
                      brushText As New SolidBrush(Color.FromArgb(20, 25, 35)),
                      brushHeaderBg As New SolidBrush(Color.FromArgb(240, 244, 248)),
                      brushCurrentBg As New SolidBrush(Color.FromArgb(232, 244, 253)),
                      brushAccent As New SolidBrush(Color.FromArgb(13, 71, 161)),
                      sfCenter As New StringFormat With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center},
                      sfLeft As New StringFormat With {.Alignment = StringAlignment.Near, .LineAlignment = StringAlignment.Center}

                    DrawOfficialHeader(g, bounds, currentY, fontHeaderSmall, fontBold, fontHeaderDoc, brushText, penThick, sfCenter)
                    DrawMetadataBox(g, bounds, currentY, fontBold, fontHeaderMain, fontRegular, brushText, brushHeaderBg, penBorder)
                    DrawMilestoneRoadmap(g, bounds, currentY, fontSectionTitle, fontBold, fontRegular, fontSmall, brushText, brushHeaderBg, brushCurrentBg, brushAccent, penBorder, penDotted, sfCenter, sfLeft)
                    DrawDirectivesAndPunchlist(g, bounds, currentY, fontSectionTitle, fontRegular, brushText, penBorder)
                    DrawHandlingNotice(g, bounds, currentY, fontSmall, brushText, brushHeaderBg, penBorder)
                End Using
            Finally
                g.Restore(pageState)
            End Try
            e.HasMorePages = False
        End Sub

        Private Sub DrawOfficialHeader(g As Graphics, bounds As RectangleF, ByRef currentY As Single, fontSmall As Font, fontBold As Font, fontDoc As Font, brushText As Brush, penThick As Pen, sfCenter As StringFormat)
            Dim strGov = "BANGSAMORO AUTONOMOUS REGION IN MUSLIM MINDANAO"
            Dim strOffice = "BANGSAMORO TRANSITION AUTHORITY  |  OFFICE OF THE SECRETARY-GENERAL"
            Dim strDocTitle = "OFFICIAL DOCUMENT ROUTING & ACTION SLIP"

            g.DrawString(strGov, fontSmall, brushText, New RectangleF(bounds.Left, currentY, bounds.Width, 16), sfCenter)
            currentY += 18
            g.DrawString(strOffice, fontBold, brushText, New RectangleF(bounds.Left, currentY, bounds.Width, 18), sfCenter)
            currentY += 20
            g.DrawString(strDocTitle, fontDoc, brushText, New RectangleF(bounds.Left, currentY, bounds.Width, 24), sfCenter)
            currentY += 28

            g.DrawLine(penThick, bounds.Left, currentY, bounds.Right, currentY)
            currentY += 6
        End Sub

        Private Sub DrawMetadataBox(g As Graphics, bounds As RectangleF, ByRef currentY As Single, fontBold As Font, fontMain As Font, fontRegular As Font, brushText As Brush, brushHeaderBg As Brush, penBorder As Pen)
            Dim metaBoxHeight As Single = 105
            Dim rectMeta As New RectangleF(bounds.Left, currentY, bounds.Width, metaBoxHeight)
            g.FillRectangle(brushHeaderBg, rectMeta)
            g.DrawRectangle(penBorder, rectMeta.X, rectMeta.Y, rectMeta.Width, rectMeta.Height)

            Dim colW As Single = bounds.Width / 2.0F
            Dim rowH As Single = 18.0F
            Dim padX As Single = 8.0F
            Dim innerY As Single = currentY + 6

            ' Left Column
            g.DrawString("Document Code:", fontBold, brushText, bounds.Left + padX, innerY)
            g.DrawString(_doc.DocCode, fontMain, brushText, bounds.Left + padX + 115, innerY - 2)
            innerY += 22

            g.DrawString("Classification:", fontBold, brushText, bounds.Left + padX, innerY)
            g.DrawString(GetDocTypeDisplay(_doc), fontRegular, brushText, bounds.Left + padX + 115, innerY)
            innerY += rowH

            g.DrawString("Flow Direction:", fontBold, brushText, bounds.Left + padX, innerY)
            g.DrawString(_doc.FlowDirection, fontRegular, brushText, bounds.Left + padX + 115, innerY)
            innerY += rowH

            g.DrawString("Document Title:", fontBold, brushText, bounds.Left + padX, innerY)
            Dim rectTitle As New RectangleF(bounds.Left + padX + 115, innerY, colW - padX - 120, 36)
            g.DrawString(_doc.Title, fontRegular, brushText, rectTitle)

            ' Right Column
            innerY = currentY + 6
            Dim rightColX As Single = bounds.Left + colW + padX

            g.DrawString("Date Registered:", fontBold, brushText, rightColX, innerY)
            g.DrawString(_doc.RegisteredAtUTC.ToString("yyyy-MM-dd HH:mm UTC"), fontRegular, brushText, rightColX + 120, innerY)
            innerY += rowH

            g.DrawString("Target Deadline:", fontBold, brushText, rightColX, innerY)
            Dim deadlineStr As String = If(_doc.TargetDeadlineUTC.HasValue, _doc.TargetDeadlineUTC.Value.ToString("yyyy-MM-dd HH:mm UTC"), "None Specified")
            g.DrawString(deadlineStr, fontRegular, brushText, rightColX + 120, innerY)
            innerY += rowH

            g.DrawString("Origin Office:", fontBold, brushText, rightColX, innerY)
            g.DrawString(If(String.IsNullOrEmpty(_doc.OriginOffice), "N/A", _doc.OriginOffice), fontRegular, brushText, rightColX + 120, innerY)
            innerY += rowH

            g.DrawString("Destination Office:", fontBold, brushText, rightColX, innerY)
            g.DrawString(If(String.IsNullOrEmpty(_doc.DestinationOffice), "N/A", _doc.DestinationOffice), fontRegular, brushText, rightColX + 120, innerY)
            innerY += rowH

            Dim landmark As String = ""
            If Not String.IsNullOrWhiteSpace(_doc.CabinetID) Then landmark = "Cabinet " & _doc.CabinetID
            If Not String.IsNullOrWhiteSpace(_doc.ShelfNo) Then landmark &= If(landmark = "", "Shelf ", "  |  Shelf ") & _doc.ShelfNo
            If Not String.IsNullOrWhiteSpace(_doc.BoxCode) Then landmark &= If(landmark = "", "Box ", "  |  Box ") & _doc.BoxCode
            If landmark <> "" Then
                g.DrawString("Storage Location:", fontBold, brushText, rightColX, innerY)
                g.DrawString(landmark, fontRegular, brushText, rightColX + 120, innerY)
            End If

            currentY += metaBoxHeight + 12
        End Sub

        Private Sub DrawMilestoneRoadmap(g As Graphics, bounds As RectangleF, ByRef currentY As Single, fontTitle As Font, fontBold As Font, fontRegular As Font, fontSmall As Font, brushText As Brush, brushHeaderBg As Brush, brushCurrentBg As Brush, brushAccent As Brush, penBorder As Pen, penDotted As Pen, sfCenter As StringFormat, sfLeft As StringFormat)
            g.DrawString("STEP-BY-STEP WORKFLOW & TRACKING ROADMAP", fontTitle, brushText, bounds.Left, currentY)
            currentY += 18

            Dim colStepW As Single = 45
            Dim colStageW As Single = 150
            Dim colOfficeW As Single = 145
            Dim colActionW As Single = 185
            Dim colStatusW As Single = 95
            Dim colSignW As Single = bounds.Width - (colStepW + colStageW + colOfficeW + colActionW + colStatusW)

            Dim tableHeaderH As Single = 24
            Dim rectTblHdr As New RectangleF(bounds.Left, currentY, bounds.Width, tableHeaderH)
            g.FillRectangle(brushHeaderBg, rectTblHdr)
            g.DrawRectangle(penBorder, rectTblHdr.X, rectTblHdr.Y, rectTblHdr.Width, rectTblHdr.Height)

            Dim curX As Single = bounds.Left

            g.DrawString("Step", fontBold, brushText, New RectangleF(curX, currentY, colStepW, tableHeaderH), sfCenter)
            curX += colStepW
            g.DrawString("Stage / Station", fontBold, brushText, New RectangleF(curX + 4, currentY, colStageW - 8, tableHeaderH), sfLeft)
            curX += colStageW
            g.DrawString("Designated Desk", fontBold, brushText, New RectangleF(curX + 4, currentY, colOfficeW - 8, tableHeaderH), sfLeft)
            curX += colOfficeW
            g.DrawString("Action Required", fontBold, brushText, New RectangleF(curX + 4, currentY, colActionW - 8, tableHeaderH), sfLeft)
            curX += colActionW
            g.DrawString("Status", fontBold, brushText, New RectangleF(curX + 4, currentY, colStatusW - 8, tableHeaderH), sfLeft)
            curX += colStatusW
            g.DrawString("Date & Signature", fontBold, brushText, New RectangleF(curX + 4, currentY, colSignW - 8, tableHeaderH), sfLeft)

            currentY += tableHeaderH

            Dim rowHeight As Single = 50.0F
            For Each stp In _steps
                Dim rowRect As New RectangleF(bounds.Left, currentY, bounds.Width, rowHeight)

                If stp.StepStatus = "CURRENT" Then
                    g.FillRectangle(brushCurrentBg, rowRect)
                End If
                g.DrawRectangle(penBorder, rowRect.X, rowRect.Y, rowRect.Width, rowRect.Height)

                curX = bounds.Left
                g.DrawString(stp.StepNumber.ToString(), fontBold, brushText, New RectangleF(curX, currentY + 4, colStepW, rowHeight - 8), sfCenter)
                curX += colStepW

                Dim fStage = If(stp.StepStatus = "CURRENT", fontBold, fontRegular)
                g.DrawString(stp.StageName, fStage, brushText, New RectangleF(curX + 4, currentY + 4, colStageW - 8, rowHeight - 8))
                curX += colStageW

                g.DrawString(stp.ResponsibleOffice, fontRegular, brushText, New RectangleF(curX + 4, currentY + 4, colOfficeW - 8, rowHeight - 8))
                curX += colOfficeW

                g.DrawString(stp.ActionRequired, fontSmall, brushText, New RectangleF(curX + 4, currentY + 4, colActionW - 8, rowHeight - 8))
                curX += colActionW

                If stp.StepStatus = "COMPLETED" Then
                    g.DrawString("[X] COMPLETED", fontBold, brushText, New RectangleF(curX + 4, currentY + 6, colStatusW - 8, 16))
                    If stp.CompletedAtUTC.HasValue Then
                        g.DrawString(stp.CompletedAtUTC.Value.ToString("MM/dd HH:mm"), fontSmall, brushText, New RectangleF(curX + 4, currentY + 22, colStatusW - 8, 14))
                    End If
                ElseIf stp.StepStatus = "CURRENT" Then
                    g.DrawString(">>> CURRENT <<<", fontBold, brushAccent, New RectangleF(curX + 4, currentY + 6, colStatusW - 8, 16))
                    g.DrawString("In Progress", fontSmall, brushAccent, New RectangleF(curX + 4, currentY + 22, colStatusW - 8, 14))
                Else
                    g.DrawString("[ ] UPCOMING", fontRegular, brushText, New RectangleF(curX + 4, currentY + 8, colStatusW - 8, 16))
                End If
                curX += colStatusW

                If stp.StepStatus = "COMPLETED" Then
                    Dim signText = If(String.IsNullOrEmpty(stp.CompletedBy), "Signed / Verified", stp.CompletedBy)
                    g.DrawString(signText, fontSmall, brushText, New RectangleF(curX + 4, currentY + 12, colSignW - 8, 16))
                    g.DrawLine(penDotted, curX + 4, currentY + 36, curX + colSignW - 8, currentY + 36)
                ElseIf stp.StepStatus = "CURRENT" Then
                    g.DrawString("Sig: ____________________", fontSmall, brushText, curX + 4, currentY + 10)
                    g.DrawString("Date: ___________________", fontSmall, brushText, curX + 4, currentY + 28)
                Else
                    g.DrawString("Sig: ____________________", fontSmall, brushText, curX + 4, currentY + 10)
                    g.DrawString("Date: ___________________", fontSmall, brushText, curX + 4, currentY + 28)
                End If

                currentY += rowHeight
            Next

            currentY += 12
        End Sub

        Private Sub DrawDirectivesAndPunchlist(g As Graphics, bounds As RectangleF, ByRef currentY As Single, fontTitle As Font, fontRegular As Font, brushText As Brush, penBorder As Pen)
            g.DrawString("SECRETARY-GENERAL DIRECTIVES & SPECIAL INSTRUCTIONS", fontTitle, brushText, bounds.Left, currentY)
            currentY += 18

            Dim rectDir As New RectangleF(bounds.Left, currentY, bounds.Width, 65)
            g.DrawRectangle(penBorder, rectDir.X, rectDir.Y, rectDir.Width, rectDir.Height)

            Dim dirText As String = ""
            If _directives IsNot Nothing AndAlso _directives.Count > 0 Then
                For Each d In _directives
                    dirText &= $"[{d.IssuedAtUTC:yyyy-MM-dd}] Directive #{d.DirectiveTypeID}: {d.DirectiveText} "
                Next
            ElseIf Not String.IsNullOrWhiteSpace(_doc.RevisionPunchlist) Then
                dirText = "Revision Punchlist: " & _doc.RevisionPunchlist
            Else
                dirText = "Standard processing applies. Ensure full document custody compliance and timely routing."
            End If

            ' The box is a fixed 52px text region: GDI truncates with an explicit ellipsis
            ' (and never paints a line that only half-fits) instead of silently clipping
            ' mid-word on the printed slip. The full text is always in the system record.
            Dim textRect As New RectangleF(bounds.Left + 8, currentY + 6, bounds.Width - 16, 52)
            Dim sfClip As New StringFormat(StringFormat.GenericDefault) With {
                .Trimming = StringTrimming.EllipsisWord,
                .FormatFlags = StringFormatFlags.LineLimit
            }
            g.DrawString(dirText, fontRegular, brushText, textRect, sfClip)
            currentY += 75
        End Sub

        Private Sub DrawHandlingNotice(g As Graphics, bounds As RectangleF, ByRef currentY As Single, fontSmall As Font, brushText As Brush, brushHeaderBg As Brush, penBorder As Pen)
            Dim rectNotice As New RectangleF(bounds.Left, currentY, bounds.Width, 42)
            g.FillRectangle(brushHeaderBg, rectNotice)
            g.DrawRectangle(penBorder, rectNotice.X, rectNotice.Y, rectNotice.Width, rectNotice.Height)

            Dim noticeNotice = "MANDATORY CUSTODY INSTRUCTION: This routing slip must stay firmly affixed to the document folder at all times. Every handling officer or clerk must record their initials and date before forwarding to the next stage. Return completed folder to Records Section upon final action."
            g.DrawString(noticeNotice, fontSmall, brushText, New RectangleF(bounds.Left + 8, currentY + 6, bounds.Width - 16, 30))
        End Sub

        Private Function GetDocTypeDisplay(doc As Document) As String
            Select Case doc.DocumentTypeID
                Case 1 : Return "Regular Communication"
                Case 2 : Return "Legislative Document"
                Case 3 : Return "Finance / Disbursement"
                Case 4 : Return "Travel Order"
                Case Else
                    If Not String.IsNullOrEmpty(doc.DocCode) Then
                        If doc.DocCode.StartsWith("FIN") Then Return "Finance / Disbursement"
                        If doc.DocCode.StartsWith("TO") Then Return "Travel Order"
                        If doc.DocCode.StartsWith("LEG") Then Return "Legislative Document"
                        If doc.DocCode.StartsWith("COMM") Then Return "Regular Communication"
                    End If
                    Return "Standard Communication"
            End Select
        End Function
    End Class

    ' Replacement for the stock PrintPreviewDialog, whose toolbar and preview pane ignore
    ' input on the office workstations. Every control here is wired in this file: zoom
    ' buttons and Ctrl+wheel change the magnification, Print opens the system print dialog,
    ' Escape closes, and the whole page is always visible so no scrolling is needed.
    Friend Class RoutingSlipPreviewForm
        Inherits Form

        Private ReadOnly _printDoc As PrintDocument
        Private ReadOnly _preview As PrintPreviewControl
        Private ReadOnly _lblZoom As Label
        Private ReadOnly _lblPrintStatus As Label

        Public Sub New(printDoc As PrintDocument)
            _printDoc = printDoc

            AppAssets.ApplyFormIcon(Me)
            Me.Font = CivicCalmTheme.FontBody
            Me.BackColor = CivicCalmTheme.ColorCanvas
            Me.Size = New Size(1000, 800)
            Me.MinimumSize = New Size(760, 560)
            Me.StartPosition = FormStartPosition.CenterParent
            Me.KeyPreview = True

            Dim flwToolbar As New FlowLayoutPanel With {
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .FlowDirection = FlowDirection.LeftToRight,
                .WrapContents = True,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(10, 8, 10, 8)
            }

            Dim btnZoomOut As New ToolbarButton("Zoom &Out")
            AddHandler btnZoomOut.Click, Sub() ApplyZoom(_preview.Zoom - 0.25)
            flwToolbar.Controls.Add(btnZoomOut)

            _lblZoom = New Label With {
                .Text = "Fit Width",
                .AutoSize = True,
                .Font = CivicCalmTheme.FontFieldLabel,
                .ForeColor = CivicCalmTheme.ColorInk,
                .Margin = New Padding(8, 8, 8, 0)
            }
            flwToolbar.Controls.Add(_lblZoom)

            Dim btnZoomIn As New ToolbarButton("Zoom &In")
            AddHandler btnZoomIn.Click, Sub() ApplyZoom(_preview.Zoom + 0.25)
            flwToolbar.Controls.Add(btnZoomIn)

            Dim btnFit As New ToolbarButton("&Fit Width")
            AddHandler btnFit.Click, Sub() ApplyFitWidth()
            flwToolbar.Controls.Add(btnFit)

            Dim btnPrint As New Button With {
                .Text = "&Print",
                .Size = New Size(110, 30),
                .BackColor = CivicCalmTheme.ColorPrimary,
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(16, 2, 6, 2)
            }
            btnPrint.FlatAppearance.BorderSize = 0
            AddHandler btnPrint.Click, AddressOf OnPrintClicked
            flwToolbar.Controls.Add(btnPrint)

            Dim btnClose As New ToolbarButton("C&lose")
            AddHandler btnClose.Click, Sub() Me.Close()
            flwToolbar.Controls.Add(btnClose)

            _lblPrintStatus = New Label With {
                .Text = "",
                .AutoSize = True,
                .Font = CivicCalmTheme.FontMicrocopy,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .Margin = New Padding(8, 9, 0, 0)
            }
            flwToolbar.Controls.Add(_lblPrintStatus)

            _preview = New PrintPreviewControl With {
                .Dock = DockStyle.Fill,
                .Document = _printDoc,
                .AutoZoom = True,
                .BackColor = CivicCalmTheme.ColorWell
            }
            AddHandler _preview.MouseWheel, AddressOf OnPreviewWheel

            Me.Controls.Add(_preview)
            Me.Controls.Add(flwToolbar)
            Me.CancelButton = btnClose

            AddHandler Me.KeyDown, Sub(s As Object, e As KeyEventArgs)
                                       If e.Control AndAlso e.KeyCode = Keys.P Then
                                           OnPrintClicked(Me, EventArgs.Empty)
                                           e.Handled = True
                                       ElseIf e.KeyCode = Keys.Oemplus OrElse e.KeyCode = Keys.Add Then
                                           ApplyZoom(_preview.Zoom + 0.25)
                                           e.Handled = True
                                       ElseIf e.KeyCode = Keys.OemMinus OrElse e.KeyCode = Keys.Subtract Then
                                           ApplyZoom(_preview.Zoom - 0.25)
                                           e.Handled = True
                                       End If
                                   End Sub
        End Sub

        Private NotInheritable Class ToolbarButton
            Inherits Button

            Public Sub New(text As String)
                Text = text
                Size = New Size(96, 30)
                BackColor = CivicCalmTheme.ColorWell
                ForeColor = CivicCalmTheme.ColorInk
                FlatStyle = FlatStyle.Flat
                Font = CivicCalmTheme.FontFieldLabel
                Cursor = Cursors.Hand
                Margin = New Padding(0, 2, 6, 2)
                FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder
            End Sub
        End Class

        ' Ctrl+wheel zooms where the cursor is; the app-wide WheelScroller keeps handling plain
        ' wheel events, which page through the (single-page) slip without effect.
        Private Sub OnPreviewWheel(sender As Object, e As MouseEventArgs)
            If (ModifierKeys And Keys.Control) = Keys.None Then Return
            Dim notches As Integer = Math.Max(1, Math.Abs(e.Delta) \ 120)
            ApplyZoom(_preview.Zoom + (CDbl(Math.Sign(e.Delta)) * CDbl(notches) * 0.25))
        End Sub

        Private Sub ApplyZoom(value As Double)
            _preview.AutoZoom = False
            _preview.Zoom = Math.Max(0.25, Math.Min(4.0, value))
            _lblZoom.Text = CInt(_preview.Zoom * 100).ToString() & "%"
        End Sub

        Private Sub ApplyFitWidth()
            _preview.AutoZoom = True
            _lblZoom.Text = "Fit Width"
        End Sub

        Private Sub OnPrintClicked(sender As Object, e As EventArgs)
            Using printDlg As New PrintDialog()
                printDlg.Document = _printDoc
                printDlg.UseEXDialog = True
                If printDlg.ShowDialog(Me) <> DialogResult.OK Then Return
                ' Applied after the dialog: assigning PrinterSettings rebuilds
                ' DefaultPageSettings from the chosen device's devmode and discards
                ' anything set beforehand.
                RoutingSlipPrintService.ConfigurePage(_printDoc)
                Try
                    _printDoc.Print()
                    _lblPrintStatus.Text = "Slip sent to " & printDlg.PrinterSettings.PrinterName & "."
                Catch ex As Exception
                    _lblPrintStatus.Text = ""
                    MessageBox.Show("Printing failed: " & ex.Message, "Print Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End Using
        End Sub
    End Class
End Namespace
