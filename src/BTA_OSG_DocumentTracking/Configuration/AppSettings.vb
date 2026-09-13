Imports System
Imports System.IO
Imports System.Text.Json

Namespace BTA_OSG
    Public Class AppSettingsRoot
        Public Property Database As DatabaseSettings
        Public Property Rfid As RfidSettings
        Public Property PdfLink As PdfLinkSettings
        Public Property Session As SessionSettings
    End Class

    Public Class AppSettings
        Private Shared _instance As AppSettings
        Private Shared ReadOnly _lock As New Object()

        Public Property DatabaseSettings As DatabaseSettings
        Public Property RfidSettings As RfidSettings
        Public Property PdfLinkSettings As PdfLinkSettings
        Public Property SessionSettings As SessionSettings

        Private Sub New()
            DatabaseSettings = New DatabaseSettings()
            RfidSettings = New RfidSettings()
            PdfLinkSettings = New PdfLinkSettings()
            SessionSettings = New SessionSettings()
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
                Dim configPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "appsettings.json")
                If File.Exists(configPath) Then
                    Dim json As String = File.ReadAllText(configPath)
                    Dim options As New JsonSerializerOptions With {
                        .PropertyNameCaseInsensitive = True
                    }
                    Dim root As AppSettingsRoot = JsonSerializer.Deserialize(Of AppSettingsRoot)(json, options)
                    If root IsNot Nothing Then
                        If root.Database IsNot Nothing Then settings.DatabaseSettings = root.Database
                        If root.Rfid IsNot Nothing Then settings.RfidSettings = root.Rfid
                        If root.PdfLink IsNot Nothing Then settings.PdfLinkSettings = root.PdfLink
                        If root.Session IsNot Nothing Then settings.SessionSettings = root.Session
                    End If
                End If
            Catch ex As Exception
                ' Silently fall back to defaults if error reading or parsing
            End Try
            Return settings
        End Function
    End Class
End Namespace
