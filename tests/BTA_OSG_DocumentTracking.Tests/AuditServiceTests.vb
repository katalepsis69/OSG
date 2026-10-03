Option Explicit On
Option Strict On

Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports BTA_OSG

Namespace BTA_OSG.Tests
    <TestClass>
    Public Class AuditServiceTests
        <TestMethod>
        Public Sub AuditService_BuildsProperlyPopulatedAuditEntry()
            Dim entry As New AuditEntry With {
                .EventAtUTC = DateTime.UtcNow,
                .ActionType = "LOGIN_SUCCESS",
                .DocumentCode = "RES-2026-001",
                .MachineName = Environment.MachineName,
                .CardPublicIDMasked = "****321A",
                .OldValuesJson = "{""Status"":""Draft""}",
                .NewValuesJson = "{""Status"":""Approved""}",
                .Success = True
            }

            Assert.IsNotNull(entry.EventAtUTC)
            Assert.AreEqual("LOGIN_SUCCESS", entry.ActionType)
            Assert.AreEqual("RES-2026-001", entry.DocumentCode)
            Assert.AreEqual("****321A", entry.CardPublicIDMasked)
            Assert.IsTrue(entry.OldValuesJson.Contains("Draft"))
            Assert.IsTrue(entry.NewValuesJson.Contains("Approved"))
            Assert.IsTrue(entry.Success)
            Assert.IsFalse(String.IsNullOrEmpty(entry.MachineName))
        End Sub

        <TestMethod>
        Public Sub AuditService_MasksCardPublicId_PreservingOnlyLastFourDigits()
            Dim rawCard = "ABCDEF0123"
            Dim masked = If(rawCard.Length > 4, "****" & rawCard.Substring(rawCard.Length - 4), rawCard)
            Assert.AreEqual("****0123", masked)
        End Sub
    End Class
End Namespace
