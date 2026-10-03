Option Explicit On
Option Strict On

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

        <TestMethod>
        Public Sub EmbeddedDB_ValidateGDriveURL_ValidDriveUrl_ReturnsTrue()
            Dim err = ""
            Dim valid = EmbeddedDB.ValidateGDriveURL("https://drive.google.com/file/d/sample-123/view", err)
            Assert.IsTrue(valid)
            Assert.AreEqual("", err)
        End Sub

        <TestMethod>
        Public Sub EmbeddedDB_ValidateGDriveURL_UntrustedHost_ReturnsFalse()
            Dim err = ""
            Dim valid = EmbeddedDB.ValidateGDriveURL("https://evil-hacker.com/malware.pdf", err)
            Assert.IsFalse(valid)
            Assert.IsTrue(err.Contains("Allowed domain hosts"))
        End Sub

        <TestMethod>
        Public Sub EmbeddedDB_ValidateGDriveURL_ExistingLocalPdf_ReturnsTrue()
            Dim tempPdf = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "valid_scan_" & Guid.NewGuid().ToString("N") & ".pdf")
            System.IO.File.WriteAllText(tempPdf, "%PDF-1.4 test")
            Try
                Dim err = ""
                Dim valid = EmbeddedDB.ValidateGDriveURL(tempPdf, err)
                Assert.IsTrue(valid)
                Assert.AreEqual("", err)
            Finally
                If System.IO.File.Exists(tempPdf) Then System.IO.File.Delete(tempPdf)
            End Try
        End Sub

        <TestMethod>
        Public Sub EmbeddedDB_ValidateGDriveURL_ExistingLocalNonPdf_ReturnsFalse()
            Dim tempBat = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "malicious_" & Guid.NewGuid().ToString("N") & ".bat")
            System.IO.File.WriteAllText(tempBat, "echo exploit")
            Try
                Dim err = ""
                Dim valid = EmbeddedDB.ValidateGDriveURL(tempBat, err)
                Assert.IsFalse(valid)
                Assert.IsTrue(err.Contains("Only PDF files"))
            Finally
                If System.IO.File.Exists(tempBat) Then System.IO.File.Delete(tempBat)
            End Try
        End Sub
    End Class
End Namespace
