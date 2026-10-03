Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.Json

Namespace BTA_OSG
    Public Class AppSettingsRoot
        Public Property IsConfigured As Boolean = False
        Public Property AnalyticsEnabled As Boolean = True
        Public Property Database As DatabaseSettings
        Public Property Rfid As RfidSettings
        Public Property PdfLink As PdfLinkSettings
        Public Property Session As SessionSettings
        Public Property Portal As PortalSettings
    End Class

    Public Class AppSettings
        Private Shared _instance As AppSettings
        Private Shared ReadOnly _lock As New Object()

        Public Property IsConfigured As Boolean = False
        Public Property AnalyticsEnabled As Boolean = True
        Public Property DatabaseSettings As DatabaseSettings
        Public Property RfidSettings As RfidSettings
        Public Property PdfLinkSettings As PdfLinkSettings
        Public Property SessionSettings As SessionSettings
        Public Property PortalSettings As PortalSettings

        ' Config layers dropped while loading, so a station that silently falls back to the wrong
        ' server can be told which file it lost.
        Public Property IgnoredConfigLayers As New List(Of String)()

        Public Sub New()
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

        ' Files the app saves are DPAPI-encrypted for the current Windows user (the SQL
        ' login password rides inside the connection string). Plain files written by IT
        ' or by older builds still load, and the next Save re-encrypts them. A protected
        ' file copied to another machine or user is unreadable there by design.
        Private Const EncryptedMarker As String = "BTA-ENC1:"

        Private Shared Function ProtectJson(json As String) As String
            Dim cipher As Byte() = ProtectedData.Protect(Encoding.UTF8.GetBytes(json), Nothing, DataProtectionScope.CurrentUser)
            Return EncryptedMarker & Convert.ToBase64String(cipher)
        End Function

        Private Shared Function TryReadLayer(raw As String, ByRef json As String) As Boolean
            Dim trimmed = raw.TrimStart()
            If Not trimmed.StartsWith(EncryptedMarker, StringComparison.Ordinal) Then
                json = raw
                Return True
            End If
            Try
                json = Encoding.UTF8.GetString(ProtectedData.Unprotect(Convert.FromBase64String(trimmed.Substring(EncryptedMarker.Length)), Nothing, DataProtectionScope.CurrentUser))
                Return True
            Catch ex As Exception
                System.Diagnostics.Trace.TraceWarning($"AppSettings layer unreadable (saved by another user or machine?): {ex.Message}")
                Return False
            End Try
        End Function

        Private Shared Function Load() As AppSettings
            Dim settings As New AppSettings()
            Try
                Dim baseDir = AppDomain.CurrentDomain.BaseDirectory
                Dim env = Environment.GetEnvironmentVariable("BTA_ENVIRONMENT")
                Dim paths As New List(Of String) From {Path.Combine(baseDir, "Resources", "appsettings.json")}
                If Not String.IsNullOrWhiteSpace(env) Then paths.Add(Path.Combine(baseDir, "Resources", $"appsettings.{env}.json"))
                ' Default to Production settings if no environment variable is set
                If String.IsNullOrWhiteSpace(env) Then paths.Add(Path.Combine(baseDir, "Resources", "appsettings.Production.json"))
                ' The AppData file is the fallback written when the install folder is read-only,
                ' so it applies before the wizard's appsettings.local.json, never after it.
                Dim appData = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "BTA_OSG_DocumentTracking", "appsettings.json")
                paths.Add(appData)
                paths.Add(Path.Combine(baseDir, "Resources", "appsettings.local.json"))

                Dim options As New JsonSerializerOptions With {.PropertyNameCaseInsensitive = True}
                For Each configPath In paths
                    If Not File.Exists(configPath) Then Continue For
                    Dim json As String = ""
                    Try
                        If Not TryReadLayer(File.ReadAllText(configPath), json) Then
                            settings.IgnoredConfigLayers.Add(Path.GetFileName(configPath) & ": could not be decrypted by this user or machine")
                            Continue For
                        End If
                        Dim root As AppSettingsRoot = JsonSerializer.Deserialize(Of AppSettingsRoot)(json, options)
                        If root Is Nothing Then Continue For
                        settings.IsConfigured = root.IsConfigured
                        settings.AnalyticsEnabled = root.AnalyticsEnabled
                        If root.Database IsNot Nothing Then settings.DatabaseSettings = root.Database
                        If root.Rfid IsNot Nothing Then settings.RfidSettings = root.Rfid
                        If root.PdfLink IsNot Nothing Then settings.PdfLinkSettings = root.PdfLink
                        If root.Session IsNot Nothing Then settings.SessionSettings = root.Session
                        If root.Portal IsNot Nothing Then settings.PortalSettings = root.Portal
                    Catch ex As Exception
                        ' One malformed layer must not discard the layers after it: a stray
                        ' backslash used to kill the loop, silently dropping the station's own
                        ' saved connection and leaving the shipped default server in place.
                        settings.IgnoredConfigLayers.Add(Path.GetFileName(configPath) & ": " & ex.Message)
                        System.Diagnostics.Trace.TraceWarning($"AppSettings layer ignored ({configPath}): {ex.Message}")
                    End Try
                Next
            Catch ex As Exception
                System.Diagnostics.Trace.TraceWarning($"AppSettings load failed: {ex.Message}")
            End Try
            Return settings
        End Function

        Public Shared Sub Save(settings As AppSettings)
            SyncLock _lock
                settings.IsConfigured = True
            Dim root As New AppSettingsRoot With {
                .IsConfigured = True,
                .AnalyticsEnabled = settings.AnalyticsEnabled,
                .Database = settings.DatabaseSettings,
                    .Rfid = settings.RfidSettings,
                    .PdfLink = settings.PdfLinkSettings,
                    .Session = settings.SessionSettings,
                    .Portal = settings.PortalSettings
                }
                Dim options As New JsonSerializerOptions With {.WriteIndented = True}
                Dim json = ProtectJson(JsonSerializer.Serialize(root, options))
                Dim baseDir = AppDomain.CurrentDomain.BaseDirectory
                Dim targetPath = Path.Combine(baseDir, "Resources", "appsettings.local.json")
                Try
                    Directory.CreateDirectory(Path.GetDirectoryName(targetPath))
                    File.WriteAllText(targetPath, json)
                Catch ex As Exception
                    Dim appDataDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "BTA_OSG_DocumentTracking")
                    Directory.CreateDirectory(appDataDir)
                    File.WriteAllText(Path.Combine(appDataDir, "appsettings.json"), json)
                End Try
                _instance = settings
            End SyncLock
        End Sub
    End Class
End Namespace
