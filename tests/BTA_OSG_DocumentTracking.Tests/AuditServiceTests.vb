Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports BTA_OSG

Namespace BTA_OSG.Tests
    <TestClass>
    Public Class AuditServiceTests
        <TestMethod>
        Public Sub AuditEntry_PopulatesMandatoryFields()
            Dim entry As New AuditEntry With {
                .EventAtUTC = DateTime.UtcNow,
                .ActionType = "LOGIN_SUCCESS",
                .DocumentCode = "RES-2026-001",
                .MachineName = Environment.MachineName,
                .Success = True
            }

            Assert.IsNotNull(entry.EventAtUTC)
            Assert.AreEqual("LOGIN_SUCCESS", entry.ActionType)
            Assert.AreEqual("RES-2026-001", entry.DocumentCode)
            Assert.IsTrue(entry.Success)
            Assert.IsFalse(String.IsNullOrEmpty(entry.MachineName))
        End Sub
    End Class
End Namespace
