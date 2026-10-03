Option Explicit On
Option Strict On

Namespace BTA_OSG
    Public Class PdfLinkSettings
        Public Property AllowedHosts As String() = {"drive.google.com", "docs.google.com"}
        Public Property RequireHttps As Boolean = True
    End Class
End Namespace
