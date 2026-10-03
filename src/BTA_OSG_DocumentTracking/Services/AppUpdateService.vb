Option Explicit On
Option Strict On

Imports System
Imports System.Diagnostics
Imports System.IO
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Reflection
Imports System.Text.Json
Imports System.Threading.Tasks
Imports System.Windows.Forms

Namespace BTA_OSG
    Public Class UpdateInfo
        Public Property Success As Boolean
        Public Property HasUpdate As Boolean
        Public Property LatestVersion As String = ""
        Public Property DownloadUrl As String = ""
        Public Property ReleaseNotes As String = ""
        Public Property ErrorMessage As String = ""
    End Class

    ''' <summary>
    ''' Minimal 1-click self-updater using native .NET stdlib and GitHub Releases.
    ''' </summary>
    Public Module AppUpdateService
        Private Const GitHubRepo As String = "katalepsis69/OSG"
        Private ReadOnly Http As New HttpClient()

        Public Async Function CheckUpdateInfoAsync() As Task(Of UpdateInfo)
            Dim info As New UpdateInfo()
            Try
                Http.DefaultRequestHeaders.UserAgent.Clear()
                Http.DefaultRequestHeaders.UserAgent.Add(New ProductInfoHeaderValue("BTA_OSG_Updater", "1.0"))

                Dim url = $"https://api.github.com/repos/{GitHubRepo}/releases/latest"
                Dim response = Await Http.GetAsync(url).ConfigureAwait(False)

                If Not response.IsSuccessStatusCode Then
                    info.Success = False
                    info.ErrorMessage = $"GitHub Releases returned HTTP {response.StatusCode}."
                    Return info
                End If

                Dim json = Await response.Content.ReadAsStringAsync().ConfigureAwait(False)
                Using doc = JsonDocument.Parse(json)
                    Dim root = doc.RootElement
                    If Not root.TryGetProperty("tag_name", Nothing) Then
                        info.Success = False
                        info.ErrorMessage = "Release tag name not found."
                        Return info
                    End If

                    Dim tagName = root.GetProperty("tag_name").GetString()
                    If String.IsNullOrWhiteSpace(tagName) Then
                        info.Success = False
                        info.ErrorMessage = "Empty release tag name."
                        Return info
                    End If

                    info.LatestVersion = tagName

                    If root.TryGetProperty("body", Nothing) Then
                        info.ReleaseNotes = root.GetProperty("body").GetString()
                    End If

                    Dim cleanTag = tagName.TrimStart("v"c, "V"c)
                    Dim currentVer = Assembly.GetExecutingAssembly().GetName().Version
                    If currentVer Is Nothing Then currentVer = New Version(1, 0, 0)

                    Dim latestVer As Version = Nothing
                    If Not Version.TryParse(cleanTag, latestVer) Then
                        Version.TryParse(cleanTag & ".0", latestVer)
                    End If

                    ' Locate executable asset in release
                    Dim downloadUrl As String = ""
                    If root.TryGetProperty("assets", Nothing) Then
                        For Each asset In root.GetProperty("assets").EnumerateArray()
                            Dim name = asset.GetProperty("name").GetString()
                            If name IsNot Nothing AndAlso name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase) Then
                                downloadUrl = asset.GetProperty("browser_download_url").GetString()
                                Exit For
                            End If
                        Next
                    End If
                    info.DownloadUrl = downloadUrl

                    info.Success = True
                    info.HasUpdate = (latestVer IsNot Nothing AndAlso latestVer > currentVer AndAlso Not String.IsNullOrEmpty(downloadUrl))
                    Return info
                End Using
            Catch ex As Exception
                info.Success = False
                info.ErrorMessage = ex.Message
                Return info
            End Try
        End Function

        Public Async Function DownloadAndApplyAsync(downloadUrl As String, ownerForm As Form) As Task
            Try
                If String.IsNullOrWhiteSpace(downloadUrl) Then
                    Throw New InvalidOperationException("Download URL is empty.")
                End If

                Dim tempDir = Path.Combine(Path.GetTempPath(), "BTA_OSG_Update")
                Directory.CreateDirectory(tempDir)
                Dim tempExe = Path.Combine(tempDir, "BTA_OSG_DocumentTracking.exe")

                Http.DefaultRequestHeaders.UserAgent.Clear()
                Http.DefaultRequestHeaders.UserAgent.Add(New ProductInfoHeaderValue("BTA_OSG_Updater", "1.0"))

                Dim exeBytes = Await Http.GetByteArrayAsync(downloadUrl).ConfigureAwait(True)
                File.WriteAllBytes(tempExe, exeBytes)

                Dim currentExe = Environment.ProcessPath
                If String.IsNullOrEmpty(currentExe) Then
                    currentExe = Application.ExecutablePath
                End If
                Dim currentPid = Environment.ProcessId
                Dim scriptPath = Path.Combine(tempDir, "apply_update.cmd")

                Dim script = "@echo off" & vbCrLf &
                             ":wait" & vbCrLf &
                             $"tasklist /fi ""PID eq {currentPid}"" | findstr ""{currentPid}"" >nul" & vbCrLf &
                             "if not errorlevel 1 (timeout /t 1 /nobreak >nul & goto wait)" & vbCrLf &
                             $"copy /y ""{tempExe}"" ""{currentExe}"" >nul" & vbCrLf &
                             $"powershell -NoProfile -Command ""Unblock-File '{currentExe}'"" >nul 2>&1" & vbCrLf &
                             $"start """" ""{currentExe}""" & vbCrLf &
                             "(goto) 2>nul & del ""%~f0"""

                File.WriteAllText(scriptPath, script)

                Process.Start(New ProcessStartInfo With {
                    .FileName = "cmd.exe",
                    .Arguments = $"/c ""{scriptPath}""",
                    .CreateNoWindow = True,
                    .UseShellExecute = False
                })

                Application.Exit()
            Catch ex As Exception
                MessageBox.Show(ownerForm, "Failed to download and apply update: " & ex.Message, "Update Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Function

        Public Async Function CheckAndApplyUpdateAsync(ownerForm As Form, manualCheck As Boolean) As Task
            Dim info = Await CheckUpdateInfoAsync().ConfigureAwait(True)
            Using dlg As New FormWhatsNew(info.HasUpdate, info.LatestVersion, info.DownloadUrl, info.ReleaseNotes)
                dlg.ShowDialog(ownerForm)
            End Using
        End Function
    End Module
End Namespace
