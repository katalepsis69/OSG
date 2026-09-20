Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports BTA_OSG

Namespace BTA_OSG.Tests
    <TestClass>
    Public Class ServiceTypingTests
        <TestMethod>
        Public Sub DocumentModel_InstantiatesWithRequiredFields()
            Dim doc As New Document With {
                .DocumentID = 101,
                .DocCode = "RES-2026-001",
                .Title = "Resolution on Parliamentary Affairs",
                .DocumentTypeID = 1,
                .StatusID = 1,
                .OriginOffice = "OSG",
                .DestinationOffice = "Plenary",
                .RegisteredByUserID = 1,
                .RegisteredAtUTC = DateTime.UtcNow
            }
            Assert.AreEqual("RES-2026-001", doc.DocCode)
            Assert.AreEqual(101, doc.DocumentID)
        End Sub

        <TestMethod>
        Public Sub ActionDirectiveModel_InstantiatesWithStrongTypes()
            Dim d As New ActionDirective With {
                .DirectiveID = 1,
                .DocumentID = 101,
                .DirectiveTypeID = 2,
                .DirectiveText = "For Immediate Review",
                .IssuedByUserID = 1,
                .IssuedAtUTC = DateTime.UtcNow,
                .IsActive = True
            }
            Assert.AreEqual("For Immediate Review", d.DirectiveText)
            Assert.IsTrue(d.IsActive)
        End Sub

        <TestMethod>
        Public Sub FormMain_InitializesWithCivicCalmColors()
            Dim form As New FormMain()
            Assert.AreEqual(CivicCalmTheme.ColorCanvas, form.BackColor)
            Assert.IsFalse(form.Text.Contains("—"))
            form.Dispose()
        End Sub
    End Class
End Namespace
