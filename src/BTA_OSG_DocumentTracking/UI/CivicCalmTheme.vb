Option Explicit On
Option Strict On

Imports System.Drawing

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
        Public Shared ReadOnly ColorAccentSG As Color = ColorTranslator.FromHtml("#B08524")
        Public Shared ReadOnly ColorAccentSoft As Color = ColorTranslator.FromHtml("#FEF9E7")
        Public Shared ReadOnly ColorDanger As Color = ColorTranslator.FromHtml("#A93226")
        Public Shared ReadOnly ColorDangerSoft As Color = ColorTranslator.FromHtml("#FBEAE8")

        ' Segoe UI Typography Ladder
        Public Shared ReadOnly FontFormTitle As New Font("Segoe UI", 12.0F, FontStyle.Bold)
        Public Shared ReadOnly FontSectionHeader As New Font("Segoe UI", 10.0F, FontStyle.Bold)
        Public Shared ReadOnly FontFieldLabel As New Font("Segoe UI", 9.0F, FontStyle.Bold)
        Public Shared ReadOnly FontBody As New Font("Segoe UI", 9.0F, FontStyle.Regular)
        Public Shared ReadOnly FontTabular As New Font("Segoe UI", 9.0F, FontStyle.Regular)
        Public Shared ReadOnly FontIdentifier As New Font("Segoe UI", 8.5F, FontStyle.Regular)
        Public Shared ReadOnly FontMicrocopy As New Font("Segoe UI", 8.0F, FontStyle.Regular)
    End Class
End Namespace
