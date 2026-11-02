Option Explicit On
Option Strict On

Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports BTA_OSG

Namespace BTA_OSG.Tests
    ''' <summary>
    ''' EmbeddedDB.ValidateGDriveURL is the one validator the launch path uses; these tests
    ''' pin its contract (https, Google hosts, local .pdf only).
    ''' </summary>
    <TestClass>
    Public Class PdfLinkValidationTests

        <TestMethod>
        Public Sub ValidateGDriveURL_ValidDriveUrl_ReturnsTrue()
            Dim err = ""
            Dim valid = EmbeddedDB.ValidateGDriveURL("https://drive.google.com/file/d/sample-123/view", err)
            Assert.IsTrue(valid)
            Assert.AreEqual("", err)
        End Sub

        <TestMethod>
        Public Sub ValidateGDriveURL_ValidDocsSubdomain_ReturnsTrue()
            Dim err = ""
            Dim valid = EmbeddedDB.ValidateGDriveURL("https://docs.google.com/document/d/sample-123/edit", err)
            Assert.IsTrue(valid)
        End Sub

        <TestMethod>
        Public Sub ValidateGDriveURL_NonHttpsLink_ReturnsFalse()
            Dim err = ""
            Dim valid = EmbeddedDB.ValidateGDriveURL("http://drive.google.com/file/d/sample-123/view", err)
            Assert.IsFalse(valid)
            Assert.IsTrue(err.Contains("https"))
        End Sub

        <TestMethod>
        Public Sub ValidateGDriveURL_UntrustedHost_ReturnsFalse()
            Dim err = ""
            Dim valid = EmbeddedDB.ValidateGDriveURL("https://evil-hacker.com/malware.pdf", err)
            Assert.IsFalse(valid)
            Assert.IsTrue(err.Contains("Allowed domain hosts"))
        End Sub

        <TestMethod>
        Public Sub ValidateGDriveURL_EmptyOrWhitespace_ReturnsTrue()
            Dim err = ""
            Assert.IsTrue(EmbeddedDB.ValidateGDriveURL("", err))
            Assert.IsTrue(EmbeddedDB.ValidateGDriveURL("   ", err))
        End Sub

        <TestMethod>
        Public Sub ValidateGDriveURL_ExistingLocalPdf_ReturnsTrue()
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
        Public Sub ValidateGDriveURL_ExistingLocalNonPdf_ReturnsFalse()
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

        <TestMethod>
        Public Sub ValidateGDriveURL_MalformedUrl_ReturnsFalse()
            Dim err = ""
            Dim valid = EmbeddedDB.ValidateGDriveURL("https://", err)
            Assert.IsFalse(valid)
        End Sub
    End Class
End Namespace
