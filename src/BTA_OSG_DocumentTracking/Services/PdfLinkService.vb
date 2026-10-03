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

        Public Function ValidateUrl(url As String, ByRef errorMessage As String) As Boolean
            errorMessage = ""
            If String.IsNullOrWhiteSpace(url) Then Return True
            
            Try
                Dim uri As New Uri(url.Trim())
                If _settings.RequireHttps AndAlso uri.Scheme <> Uri.UriSchemeHttps Then
                    errorMessage = "HTTPS is required."
                    Return False
                End If
                Dim host = uri.Host.ToLowerInvariant()
                Dim isAllowed = False
                If _settings.AllowedHosts IsNot Nothing Then
                    For Each allowedHost In _settings.AllowedHosts
                        Dim ah = allowedHost.ToLowerInvariant().Trim()
                        If host = ah Then
                            isAllowed = True
                            Exit For
                        End If
                    Next
                End If

                If Not isAllowed Then
                    errorMessage = "Host not allowed. Restricted to Google Drive / Docs."
                    Return False
                End If
                Return True
            Catch ex As UriFormatException
                errorMessage = "Invalid URL."
                Return False
            End Try
        End Function

        Public Sub LaunchPdf(url As String, userId As Integer)
            Dim err = ""
            If Not ValidateUrl(url, err) Then Throw New Exception(err)
            
            If Not String.IsNullOrWhiteSpace(url) Then
                Process.Start(New ProcessStartInfo(url) With { .UseShellExecute = True })
                If _auditService IsNot Nothing Then
                    _auditService.LogEvent("PDF_LINK_OPENED", "PDF", Nothing, url, Nothing, Nothing, True, Nothing)
                End If
            End If
        End Sub
    End Class
End Namespace
