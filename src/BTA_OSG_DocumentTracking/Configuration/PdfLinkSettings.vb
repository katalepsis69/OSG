Option Explicit On
Option Strict On

Namespace BTA_OSG
    Public Class PdfLinkSettings
        ' The Google Drive for Desktop sync folder that holds office scans. When set, an
        ' attached scan is copied here so Drive uploads it in the background and every desk
        ' syncing the same folder previews the same stored path. Empty disables staging.
        ' URL host allowlisting lives in EmbeddedDB.ValidateGDriveURL, the one validator
        ' the launch path actually uses.
        Public Property ScansFolder As String = ""
    End Class
End Namespace
