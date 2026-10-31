Option Explicit On
Option Strict On

Imports System.Drawing
Imports System.Windows.Forms

Namespace BTA_OSG
    Public NotInheritable Class CivicCalmTheme
        Private Sub New()
        End Sub

        ' 3-Tier Optical Surface Hierarchy
        Public Shared ReadOnly ColorCanvas As Color = ColorTranslator.FromHtml("#F4F6F8")
        Public Shared ReadOnly ColorSurface As Color = ColorTranslator.FromHtml("#FFFFFF")
        Public Shared ReadOnly ColorWell As Color = ColorTranslator.FromHtml("#EEF2F5")
        Public Shared ReadOnly ColorBorder As Color = ColorTranslator.FromHtml("#DDE2E5")

        ' Typography & Ink Tokens
        Public Shared ReadOnly ColorInk As Color = ColorTranslator.FromHtml("#1B242C")
        Public Shared ReadOnly ColorInkMuted As Color = ColorTranslator.FromHtml("#55606A")

        ' Action & Semantic Tokens
        Public Shared ReadOnly ColorPrimary As Color = ColorTranslator.FromHtml("#146A3D")
        Public Shared ReadOnly ColorPrimarySoft As Color = ColorTranslator.FromHtml("#E6F2EB")
        Public Shared ReadOnly ColorAccentSG As Color = ColorTranslator.FromHtml("#8C6414")
        Public Shared ReadOnly ColorAccentSoft As Color = ColorTranslator.FromHtml("#FEF9E7")
        Public Shared ReadOnly ColorDanger As Color = ColorTranslator.FromHtml("#A93226")
        Public Shared ReadOnly ColorDangerSoft As Color = ColorTranslator.FromHtml("#FBEAE8")

        ' Status & Semantic Badges
        Public Shared ReadOnly ColorStatusRevisionBg As Color = ColorTranslator.FromHtml("#FEE2E2")
        Public Shared ReadOnly ColorStatusRevisionFg As Color = ColorTranslator.FromHtml("#991B1B")
        Public Shared ReadOnly ColorStatusReviewBg As Color = ColorTranslator.FromHtml("#E0F2FE")
        Public Shared ReadOnly ColorStatusReviewFg As Color = ColorTranslator.FromHtml("#0369A1")
        Public Shared ReadOnly ColorStatusApprovedBg As Color = ColorTranslator.FromHtml("#D1FAE5")
        Public Shared ReadOnly ColorStatusApprovedFg As Color = ColorTranslator.FromHtml("#065F46")
        Public Shared ReadOnly ColorStatusReleasedBg As Color = ColorTranslator.FromHtml("#E2E8F0")
        Public Shared ReadOnly ColorStatusReleasedFg As Color = ColorTranslator.FromHtml("#1E293B")
        Public Shared ReadOnly ColorStatusReceivedBg As Color = ColorTranslator.FromHtml("#FEF3C7")
        Public Shared ReadOnly ColorStatusReceivedFg As Color = ColorTranslator.FromHtml("#92400E")
        Public Shared ReadOnly ColorStatusPendingBg As Color = ColorTranslator.FromHtml("#ECEDF0")
        Public Shared ReadOnly ColorStatusPendingFg As Color = ColorTranslator.FromHtml("#52606B")
        Public Shared ReadOnly ColorStatusArchivedBg As Color = ColorTranslator.FromHtml("#E2E6EA")
        Public Shared ReadOnly ColorStatusArchivedFg As Color = ColorTranslator.FromHtml("#3A4550")

        ' Punchlist & Warning Alert Accents
        Public Shared ReadOnly ColorPunchlistBorder As Color = ColorTranslator.FromHtml("#FCA5A5")
        Public Shared ReadOnly ColorPunchlistContentFg As Color = ColorTranslator.FromHtml("#7F1D1D")
        Public Shared ReadOnly ColorInfo As Color = ColorTranslator.FromHtml("#0369A1")
        Public Shared ReadOnly ColorWarning As Color = ColorTranslator.FromHtml("#B45309")

        ' Segoe UI Typography Ladder
        Public Shared ReadOnly FontFormTitle As New Font("Segoe UI", 12.0F, FontStyle.Bold)
        Public Shared ReadOnly FontSectionHeader As New Font("Segoe UI", 10.0F, FontStyle.Bold)
        Public Shared ReadOnly FontFieldLabel As New Font("Segoe UI", 9.0F, FontStyle.Bold)
        Public Shared ReadOnly FontBody As New Font("Segoe UI", 9.0F, FontStyle.Regular)
        Public Shared ReadOnly FontTabular As New Font("Segoe UI", 9.0F, FontStyle.Regular)
        Public Shared ReadOnly FontIdentifier As New Font("Segoe UI", 8.5F, FontStyle.Regular)
        Public Shared ReadOnly FontMicrocopy As New Font("Segoe UI", 8.0F, FontStyle.Regular)

        ''' <summary>
        ''' Central factory for consistent, accessible icon buttons across all forms.
        ''' Standardizes AutoSize, minimum height, flat style, hand cursor, padding,
        ''' and aligned icon placement.
        ''' </summary>
        Public Shared Function CreateIconButton(text As String,
                                                Optional iconName As String = Nothing,
                                                Optional iconSize As Integer = 16,
                                                Optional backColor As Color? = Nothing,
                                                Optional foreColor As Color? = Nothing,
                                                Optional minHeight As Integer = 34,
                                                Optional margin As Padding? = Nothing,
                                                Optional tabIndex As Integer? = Nothing,
                                                Optional visible As Boolean? = Nothing) As Button
            Dim bg = If(backColor.HasValue, backColor.Value, ColorWell)
            Dim fg = If(foreColor.HasValue, foreColor.Value, If(bg = ColorWell OrElse bg = ColorSurface, ColorInk, Color.White))
            Dim btn As New Button With {
                .Text = text,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .MinimumSize = New Size(0, minHeight),
                .BackColor = bg,
                .ForeColor = fg,
                .FlatStyle = FlatStyle.Flat,
                .Font = FontFieldLabel,
                .Cursor = Cursors.Hand,
                .TextAlign = ContentAlignment.MiddleCenter,
                .Padding = New Padding(6, 0, 6, 0)
            }
            If margin.HasValue Then
                btn.Margin = margin.Value
            End If
            If tabIndex.HasValue Then
                btn.TabIndex = tabIndex.Value
            End If
            If visible.HasValue Then
                btn.Visible = visible.Value
            End If
            If Not String.IsNullOrEmpty(iconName) Then
                btn.Image = AppAssets.GetIcon(iconName, iconSize, fg)
                btn.ImageAlign = ContentAlignment.MiddleLeft
                btn.TextImageRelation = TextImageRelation.ImageBeforeText
            End If
            If bg = ColorWell OrElse bg = ColorSurface Then
                btn.FlatAppearance.BorderColor = ColorBorder
            Else
                btn.FlatAppearance.BorderSize = 0
            End If
            Return btn
        End Function
    End Class
End Namespace
