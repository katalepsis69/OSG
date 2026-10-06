Option Explicit On
Option Strict On

Imports System
Imports System.Diagnostics
Imports System.Linq

Namespace BTA_OSG
    Public Class PdfLinkService
        Private ReadOnly _settings As PdfLinkSettings
        Private ReadOnly _auditService As AuditService

        Public Sub New(settings As PdfLinkSettings, auditService As AuditService)
            _settings = settings
            _auditService = auditService
        End Sub

        ''' <summary>
        ''' Files an attached scan into the configured Drive sync folder (PdfLinkSettings.ScansFolder)
        ''' under a machine-assigned SCAN- name, so Google Drive for Desktop uploads it in the
        ''' background and every desk syncing that folder previews the same stored path. With no
        ''' folder configured, or on copy failure, this reports and leaves stagedPath at the
        ''' picked path: staging is a convenience and must never block an attachment.
        ''' </summary>
        Public Function StageScan(sourcePath As String, ByRef stagedPath As String, ByRef errMessage As String) As Boolean
            stagedPath = sourcePath
            errMessage = ""
            If String.IsNullOrWhiteSpace(_settings.ScansFolder) Then Return True

            If Not System.IO.File.Exists(sourcePath) Then
                errMessage = "the picked file no longer exists"
                Return False
            End If
            If String.Compare(System.IO.Path.GetExtension(sourcePath), ".pdf", StringComparison.OrdinalIgnoreCase) <> 0 Then
                errMessage = "only PDF files (.pdf) may be staged as scans"
                Return False
            End If

            Try
                Dim root As String = System.IO.Path.GetFullPath(_settings.ScansFolder.Trim())
                If Not System.IO.Directory.Exists(root) Then
                    errMessage = "the scans folder does not exist: " & root
                    Return False
                End If
                Dim fullSource As String = System.IO.Path.GetFullPath(sourcePath)
                Dim rootPrefix As String = root.TrimEnd("\"c, "/"c) & "\"
                If fullSource.StartsWith(rootPrefix, StringComparison.OrdinalIgnoreCase) Then Return True

                Dim target As String = System.IO.Path.Combine(root, "SCAN-" & Date.Now.ToString("yyyyMMdd-HHmmss") & ".pdf")
                Dim bump As Integer = 1
                While System.IO.File.Exists(target)
                    target = System.IO.Path.Combine(root, "SCAN-" & Date.Now.ToString("yyyyMMdd-HHmmss") & "-" & bump.ToString() & ".pdf")
                    bump += 1
                End While
                System.IO.File.Copy(fullSource, target)
                stagedPath = target
                Return True
            Catch ex As Exception
                errMessage = ex.Message
                Return False
            End Try
        End Function
    End Class
End Namespace
