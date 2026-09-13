Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports BTA_OSG

Namespace BTA_OSG.Tests
    <TestClass>
    Public Class PdfLinkValidationTests
        Private _service As PdfLinkService

        <TestInitialize>
        Public Sub Setup()
            Dim settings As New PdfLinkSettings With {
                .AllowedHosts = New String() {"drive.google.com", "docs.google.com"},
                .RequireHttps = True
            }
            _service = New PdfLinkService(settings, Nothing)
        End Sub

        <TestMethod>
        Public Sub ValidateUrl_ValidGoogleDriveLink_ReturnsTrue()
            Dim err = ""
            Dim valid = _service.ValidateUrl("https://drive.google.com/file/d/sample-res-001/view", err)
            Assert.IsTrue(valid)
            Assert.AreEqual("", err)
        End Sub

        <TestMethod>
        Public Sub ValidateUrl_NonHttpsLink_ReturnsFalse()
            Dim err = ""
            Dim valid = _service.ValidateUrl("http://drive.google.com/file/d/sample-res-001/view", err)
            Assert.IsFalse(valid)
            Assert.IsTrue(err.Contains("HTTPS"))
        End Sub

        <TestMethod>
        Public Sub ValidateUrl_UntrustedDomain_ReturnsFalse()
            Dim err = ""
            Dim valid = _service.ValidateUrl("https://malicious-site.com/doc.pdf", err)
            Assert.IsFalse(valid)
            Assert.IsTrue(err.Contains("Host not allowed"))
        End Sub

        <TestMethod>
        Public Sub ValidateUrl_EmptyOrWhitespace_ReturnsTrue()
            Dim err = ""
            Dim valid = _service.ValidateUrl("", err)
            Assert.IsTrue(valid)
        End Sub
    End Class
End Namespace
