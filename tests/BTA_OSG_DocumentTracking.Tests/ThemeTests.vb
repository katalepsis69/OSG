Option Explicit On
Option Strict On

Imports System.Drawing
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports BTA_OSG

Namespace BTA_OSG.Tests
    <TestClass>
    Public Class ThemeTests
        <TestMethod>
        Public Sub ThemeTokens_MatchDesignSpecification()
            Assert.AreEqual(ColorTranslator.FromHtml("#F4F6F8"), CivicCalmTheme.ColorCanvas)
            Assert.AreEqual(ColorTranslator.FromHtml("#FFFFFF"), CivicCalmTheme.ColorSurface)
            Assert.AreEqual(ColorTranslator.FromHtml("#EEF2F5"), CivicCalmTheme.ColorWell)
            Assert.AreEqual(ColorTranslator.FromHtml("#DDE2E5"), CivicCalmTheme.ColorBorder)
            Assert.AreEqual(ColorTranslator.FromHtml("#1B242C"), CivicCalmTheme.ColorInk)
            Assert.AreEqual(ColorTranslator.FromHtml("#146A3D"), CivicCalmTheme.ColorPrimary)
            Assert.AreEqual(ColorTranslator.FromHtml("#E6F2EB"), CivicCalmTheme.ColorPrimarySoft)
        End Sub

        <TestMethod>
        Public Sub TypographyLadder_UsesSegoeUI()
            Assert.AreEqual("Segoe UI", CivicCalmTheme.FontFormTitle.FontFamily.Name)
            Assert.AreEqual(12.0F, CivicCalmTheme.FontFormTitle.Size)
            Assert.AreEqual(FontStyle.Bold, CivicCalmTheme.FontFormTitle.Style)
            Assert.AreEqual("Segoe UI", CivicCalmTheme.FontBody.FontFamily.Name)
            Assert.AreEqual(9.0F, CivicCalmTheme.FontBody.Size)
        End Sub
    End Class
End Namespace
