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
    ''' <summary>
    ''' Minimal 1-click self-updater using native .NET stdlib and GitHub Releases.
    ''' </summary>
    Public Module AppUpdateService
        Private Const GitHubRepo As String = "katalepsis69/OSG"
        Private ReadOnly Http As New HttpClient()

        Public Async Function CheckAndApplyUpdateAsync(ownerForm As Form, manualCheck As Boolean) As Task
            Try
                Http.DefaultRequestHeaders.UserAgent.Clear()
                Http.DefaultRequestHeaders.UserAgent.Add(New ProductInfoHeaderValue("BTA_OSG_Updater", "1.0"))

                Dim url = $"https://api.github.com/repos/{GitHubRepo}/releases/latest"
                Dim response = Await Http.GetAsync(url).ConfigureAwait(True)

                If Not response.IsSuccessStatusCode Then
                    If manualCheck Then
                        MessageBox.Show(ownerForm, "Could not check for updates. Make sure internet connection is active.", "Update Check", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                    Return
                End If

                Dim json = Await response.Content.ReadAsStringAsync().ConfigureAwait(True)
                Using doc = JsonDocument.Parse(json)
                    Dim root = doc.RootElement
                    If Not root.TryGetProperty("tag_name", Nothing) Then Return
                    Dim tagName = root.GetProperty("tag_name").GetString()
                    If String.IsNullOrWhiteSpace(tagName) Then Return

                    Dim cleanTag = tagName.TrimStart("v"c, "V"c)
                    Dim currentVer = Assembly.GetExecutingAssembly().GetName().Version
                    If currentVer Is Nothing Then currentVer = New Version(1, 0, 0)

                    Dim latestVer As Version = Nothing
                    If Not Version.TryParse(cleanTag, latestVer) Then
                        Version.TryParse(cleanTag & ".0", latestVer)
                    End If

                    If latestVer IsNot Nothing AndAlso latestVer <= currentVer Then
                        If manualCheck Then
                            MessageBox.Show(ownerForm, $"You are running the latest version (v{currentVer.ToString(3)}).", "Up to Date", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End If
                        Return
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

                    If String.IsNullOrEmpty(downloadUrl) Then
                        If manualCheck Then
                            MessageBox.Show(ownerForm, $"Release {tagName} found, but no updated executable asset is available.", "Update Check", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        End If
                        Return
                    End If

                    Dim result = MessageBox.Show(ownerForm,
                        $"New version {tagName} is available (Current: v{currentVer.ToString(3)})." & vbCrLf & vbCrLf &
                        "Download and install update now?",
                        "Update Available", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

                    If result <> DialogResult.Yes Then Return

                    ' Download new binary to %TEMP%
                    Dim tempDir = Path.Combine(Path.GetTempPath(), "BTA_OSG_Update")
                    Directory.CreateDirectory(tempDir)
                    Dim tempExe = Path.Combine(tempDir, "BTA_OSG_DocumentTracking.exe")

                    Dim exeBytes = Await Http.GetByteArrayAsync(downloadUrl).ConfigureAwait(True)
                    File.WriteAllBytes(tempExe, exeBytes)

                    ' Launch minimal batch helper to swap exe once process terminates
                    Dim currentExe = Environment.ProcessPath
                    Dim currentPid = Environment.ProcessId
                    Dim scriptPath = Path.Combine(tempDir, "apply_update.cmd")

                    Dim script = "@echo off" & vbCrLf &
                                 ":wait" & vbCrLf &
                                 $"tasklist /fi ""PID eq {currentPid}"" | findstr ""{currentPid}"" >nul" & vbCrLf &
                                 "if not errorlevel 1 (timeout /t 1 /nobreak >nul & goto wait)" & vbCrLf &
                                 $"copy /y ""{tempExe}"" ""{currentExe}"" >nul" & vbCrLf &
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
                End Using
            Catch ex As Exception
                If manualCheck Then
                    MessageBox.Show(ownerForm, "Update failed: " & ex.Message, "Update Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End Try
        End Function
    End Module
End Namespace
