Option Explicit On
Option Strict On

Imports System.Data
Imports System.Drawing
Imports System.Windows.Forms
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

        <TestMethod>
        Public Sub DataGridStyler_ApplyCivicStyle_SetsExpectedPropertiesAndBindsData()
            Using dgv As New DataGridView()
                DataGridStyler.ApplyCivicStyle(dgv)
                dgv.AllowUserToAddRows = False

                Assert.IsTrue(dgv.AutoGenerateColumns)
                Assert.IsFalse(dgv.EnableHeadersVisualStyles)
                Assert.AreEqual(CivicCalmTheme.ColorSurface, dgv.BackgroundColor)
                Assert.AreEqual(CivicCalmTheme.ColorBorder, dgv.GridColor)
                Assert.AreEqual(CivicCalmTheme.ColorWell, dgv.ColumnHeadersDefaultCellStyle.BackColor)

                Using dt As New DataTable()
                    dt.Columns.Add("DocumentID", GetType(Integer))
                    dt.Columns.Add("DocCode", GetType(String))
                    dt.Columns.Add("Title", GetType(String))
                    dt.Columns.Add("CabinetID", GetType(String))

                    dt.Rows.Add(1, "OSG-2026-0001", "Test Subject", "CAB-01")

                    dgv.BindingContext = New BindingContext()
                    dgv.DataSource = dt
                    DataGridStyler.FormatDocumentColumns(dgv)

                    Assert.IsTrue(dgv.Columns.Count >= 4)
                    Assert.AreEqual(1, dgv.Rows.Count)
                    Assert.AreEqual("ID", dgv.Columns("DocumentID").HeaderText)
                    Assert.AreEqual("Document Code", dgv.Columns("DocCode").HeaderText)
                    Assert.AreEqual("Document Title / Subject", dgv.Columns("Title").HeaderText)
                    Assert.IsFalse(dgv.Columns("CabinetID").Visible)
                End Using
            End Using
        End Sub
    End Class
End Namespace
