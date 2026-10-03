Option Explicit On
Option Strict On

Namespace BTA_OSG
    Public Class PortalSettings
        Public Property PortalEnabled As Boolean = False
        Public Property BaseUrl As String = "https://portal.bta-osg.gov.ph"
        Public Property BridgeKey As String = ""
        Public Property PollIntervalSeconds As Integer = 60
        Public Property TimeoutSeconds As Integer = 10
        ' Mail credentials for the bundled portal ride the encrypted settings file. Empty
        ' means the portal cannot send mail until the operator configures them; no live
        ' key ships in source.
        Public Property BrevoApiKey As String = ""
        Public Property BrevoSenderEmail As String = ""
        Public Property BrevoSenderName As String = "BTA OSG Records Section"
    End Class
End Namespace
