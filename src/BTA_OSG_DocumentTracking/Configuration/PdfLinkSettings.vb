Option Explicit On
Option Strict On

Namespace BTA_OSG
    Public Class PdfLinkSettings
        Public Property AllowedHosts As String() = {"drive.google.com", "docs.google.com"}
        Public Property RequireHttps As Boolean = True

        ' The Google Drive for Desktop sync folder that holds office scans. When set, an
        ' attached scan is copied here so Drive uploads it in the background and every desk
        ' syncing the same folder previews the same stored path. Empty disables staging.
        Public Property ScansFolder As String = ""
    End Class
End Namespace
