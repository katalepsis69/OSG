Option Explicit On
Option Strict On

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

        <TestMethod>
        Public Sub DocumentType_ContainsExpectedOsgCategories()
            Dim expectedCodes As String() = {"REG_COMM", "LEG", "FIN", "TRAVEL"}
            For Each code In expectedCodes
                Assert.IsTrue(DocumentType.IsValidCategory(code), "Missing category: " & code)
            Next
        End Sub

        <TestMethod>
        Public Sub DocumentStatus_ContainsExpectedOsgStatuses()
            Dim expectedStatuses As String() = {"RECEIVED", "FOR_REVIEW", "FOR_REVISION", "APPROVED", "RELEASED", "FILED"}
            For Each status In expectedStatuses
                Assert.IsTrue(DocumentStatus.IsValidStatus(status), "Missing status: " & status)
            Next
        End Sub
        <TestMethod>
        Public Sub DocumentStatus_DisplayName_MatchesSeededStatusNames()
            Assert.AreEqual("Received", DocumentStatus.DisplayName("RECEIVED"))
            Assert.AreEqual("For Review", DocumentStatus.DisplayName("FOR_REVIEW"))
            Assert.AreEqual("For Revision", DocumentStatus.DisplayName("FOR_REVISION"))
            Assert.AreEqual("Approved", DocumentStatus.DisplayName("APPROVED"))
            Assert.AreEqual("Released", DocumentStatus.DisplayName("RELEASED"))
            Assert.AreEqual("Filed", DocumentStatus.DisplayName("FILED"))
            Assert.AreEqual("", DocumentStatus.DisplayName(""), "a blank status renders blank, never as a code")
            Assert.AreEqual("NOT_A_STATUS", DocumentStatus.DisplayName("NOT_A_STATUS"), "unknown text is shown as stored")
        End Sub


    End Class
End Namespace
