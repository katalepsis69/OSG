Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Text.Json

Namespace BTA_OSG
    Public Class AppSettingsRoot
        Public Property Database As DatabaseSettings
        Public Property Rfid As RfidSettings
        Public Property PdfLink As PdfLinkSettings
        Public Property Session As SessionSettings
        Public Property Portal As PortalSettings
    End Class

    Public Class AppSettings
        Private Shared _instance As AppSettings
        Private Shared ReadOnly _lock As New Object()

        Public Property DatabaseSettings As DatabaseSettings
        Public Property RfidSettings As RfidSettings
        Public Property PdfLinkSettings As PdfLinkSettings
        Public Property SessionSettings As SessionSettings
        Public Property PortalSettings As PortalSettings

        Private Sub New()
            DatabaseSettings = New DatabaseSettings()
            RfidSettings = New RfidSettings()
            PdfLinkSettings = New PdfLinkSettings()
            SessionSettings = New SessionSettings()
            PortalSettings = New PortalSettings()
        End Sub

        Public Shared ReadOnly Property Instance As AppSettings
            Get
                If _instance Is Nothing Then
                    SyncLock _lock
                        If _instance Is Nothing Then
                            _instance = Load()
                        End If
                    End SyncLock
                End If
                Return _instance
            End Get
        End Property

        Public Shared Sub Reload()
            SyncLock _lock
                _instance = Load()
            End SyncLock
        End Sub

        Private Shared Function Load() As AppSettings
            Dim settings As New AppSettings()
            Try
                Dim baseDir = AppDomain.CurrentDomain.BaseDirectory
                Dim env = Environment.GetEnvironmentVariable("BTA_ENVIRONMENT")
                Dim paths As New List(Of String) From {Path.Combine(baseDir, "Resources", "appsettings.json")}
                If Not String.IsNullOrWhiteSpace(env) Then paths.Add(Path.Combine(baseDir, "Resources", $"appsettings.{env}.json"))
                ' Default to Production settings if no environment variable is set
                If String.IsNullOrWhiteSpace(env) Then paths.Add(Path.Combine(baseDir, "Resources", "appsettings.Production.json"))
                Dim options As New JsonSerializerOptions With {.PropertyNameCaseInsensitive = True}
                For Each configPath In paths
                    If Not File.Exists(configPath) Then Continue For
                    Dim json As String = File.ReadAllText(configPath)
                    Dim root As AppSettingsRoot = JsonSerializer.Deserialize(Of AppSettingsRoot)(json, options)
                    If root Is Nothing Then Continue For
                    If root.Database IsNot Nothing Then settings.DatabaseSettings = root.Database
                    If root.Rfid IsNot Nothing Then settings.RfidSettings = root.Rfid
                    If root.PdfLink IsNot Nothing Then settings.PdfLinkSettings = root.PdfLink
                    If root.Session IsNot Nothing Then settings.SessionSettings = root.Session
                    If root.Portal IsNot Nothing Then settings.PortalSettings = root.Portal
                Next
            Catch ex As Exception
                System.Diagnostics.Trace.TraceWarning($"AppSettings load failed: {ex.Message}")
            End Try
            Return settings
        End Function
    End Class
End Namespace
