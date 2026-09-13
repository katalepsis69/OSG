Imports System
Imports System.Text.RegularExpressions
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports BTA_OSG

Namespace BTA_OSG.Tests
    <TestClass>
    Public Class DocumentCodingTests
        <TestMethod>
        Public Sub DocumentCode_Format_MatchesStandardPattern()
            Dim prefix = "RES"
            Dim year = 2026
            Dim seq = 1
            Dim code = String.Format("{0}-{1}-{2:D3}", prefix, year, seq)
            
            Assert.AreEqual("RES-2026-001", code)
            Assert.IsTrue(Regex.IsMatch(code, "^[A-Z]{3}-\d{4}-\d{3,}$"))
        End Sub

        <TestMethod>
        Public Sub DocumentCode_SequentialIncrement_ExpandsPadding()
            Dim prefix = "BLL"
            Dim year = 2026
            Dim seqLarge = 1005
            Dim code = String.Format("{0}-{1}-{2:D3}", prefix, year, seqLarge)
            
            Assert.AreEqual("BLL-2026-1005", code)
            Assert.IsTrue(Regex.IsMatch(code, "^[A-Z]{3}-\d{4}-\d{3,}$"))
        End Sub
    End Class
End Namespace
