Option Explicit On
Option Strict On

Imports System
Imports System.IO
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports BTA_OSG

''' <summary>
''' Attach Scan stages the picked PDF into the configured Drive sync folder so Google
''' Drive for Desktop uploads it and every synced desk previews the same stored path.
''' With no folder configured the picked path must be stored unchanged: staging is a
''' convenience and never blocks an attachment.
''' </summary>
<TestClass>
Public Class ScanStagingTests
    Private Shared ReadOnly TempRoot As String =
        Path.Combine(Path.GetTempPath(), "bta-scanstage-" & Guid.NewGuid().ToString("N"))

    Private Shared Function MakePdf(subDir As String, name As String) As String
        Directory.CreateDirectory(subDir)
        Dim p = Path.Combine(subDir, name)
        File.WriteAllText(p, "%PDF-1.4 staging test")
        Return p
    End Function

    <ClassCleanup>
    Public Shared Sub RemoveTempRoot()
        Try
            If Directory.Exists(TempRoot) Then Directory.Delete(TempRoot, True)
        Catch
        End Try
    End Sub

    <TestMethod>
    Public Sub StageScan_CopiesIntoConfiguredFolder()
        Dim source = MakePdf(Path.Combine(TempRoot, "src"), "orig.pdf")
        Dim scans = Path.Combine(TempRoot, "scans")
        Directory.CreateDirectory(scans)
        Dim svc As New PdfLinkService(New PdfLinkSettings With {.ScansFolder = scans}, Nothing)

        Dim staged = "", err = ""
        Assert.IsTrue(svc.StageScan(source, staged, err), err)
        Assert.AreNotEqual(source, staged, "the stored path points at the staged copy")
        Assert.IsTrue(File.Exists(staged), "the scan exists inside the scans folder")
        Assert.IsTrue(Path.GetFileName(staged).StartsWith("SCAN-"), "the staged name is machine-assigned")
        Assert.IsTrue(staged.StartsWith(Path.GetFullPath(scans), StringComparison.OrdinalIgnoreCase),
                      "the staged path lives in the scans folder")
    End Sub

    <TestMethod>
    Public Sub StageScan_AlreadyInFolder_KeepsThePath()
        Dim scans = Path.Combine(TempRoot, "scans")
        Dim source = MakePdf(scans, "already-there.pdf")
        Dim svc As New PdfLinkService(New PdfLinkSettings With {.ScansFolder = scans}, Nothing)

        Dim staged = "", err = ""
        Assert.IsTrue(svc.StageScan(source, staged, err), err)
        Assert.AreEqual(source, staged, "a scan filed in the folder is stored as is, not copied onto itself")
    End Sub

    <TestMethod>
    Public Sub StageScan_WithoutConfiguredFolder_IsANoOp()
        Dim source = MakePdf(Path.Combine(TempRoot, "src2"), "plain.pdf")
        Dim svc As New PdfLinkService(New PdfLinkSettings With {.ScansFolder = ""}, Nothing)

        Dim staged = "", err = ""
        Assert.IsTrue(svc.StageScan(source, staged, err), err)
        Assert.AreEqual(source, staged, "no folder configured: the picked path is stored as is")
    End Sub

    <TestMethod>
    Public Sub StageScan_MissingFolder_ReportsAndKeepsSource()
        Dim source = MakePdf(Path.Combine(TempRoot, "src3"), "orig.pdf")
        Dim missing = Path.Combine(TempRoot, "never-created")
        Dim svc As New PdfLinkService(New PdfLinkSettings With {.ScansFolder = missing}, Nothing)

        Dim staged = "", err = ""
        Assert.IsFalse(svc.StageScan(source, staged, err))
        Assert.AreEqual(source, staged, "failure falls back to the picked path")
        Assert.IsTrue(err.Length > 0, "the reason reaches the caller for the status line")
    End Sub

    <TestMethod>
    Public Sub StageScan_RejectsNonPdf()
        Dim dir = Path.Combine(TempRoot, "src4")
        Directory.CreateDirectory(dir)
        Dim source = Path.Combine(dir, "photo.jpg")
        File.WriteAllText(source, "not a pdf")
        Dim scans = Path.Combine(TempRoot, "scans")
        Directory.CreateDirectory(scans)
        Dim svc As New PdfLinkService(New PdfLinkSettings With {.ScansFolder = scans}, Nothing)

        Dim staged = "", err = ""
        Assert.IsFalse(svc.StageScan(source, staged, err))
        Assert.AreEqual(source, staged)
    End Sub
End Class
