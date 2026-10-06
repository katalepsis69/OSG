Option Explicit On
Option Strict On

Imports System
Imports System.Diagnostics
Imports System.Globalization
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
        Public Property PublishedAt As DateTime = DateTime.MinValue
    End Class

    ''' <summary>
    ''' Minimal 1-click self-updater using native .NET stdlib and GitHub Releases.
    ''' </summary>
    Public Module AppUpdateService
        Private Const GitHubRepo As String = "katalepsis69/OSG"
        ' The office LAN is often offline with firewalls that drop rather than refuse; without
        ' a budget the manual check spins for the 100s HttpClient default.
        Private ReadOnly Http As New HttpClient() With {.Timeout = TimeSpan.FromSeconds(10)}

        Public Async Function CheckUpdateInfoAsync() As Task(Of UpdateInfo)
            Dim info As New UpdateInfo()
            Try
                Http.DefaultRequestHeaders.UserAgent.Clear()
                Http.DefaultRequestHeaders.UserAgent.Add(New ProductInfoHeaderValue("BTA_OSG_Updater", "1.0"))

                Dim token = Environment.GetEnvironmentVariable("GITHUB_TOKEN")
                If Not String.IsNullOrEmpty(token) Then
                    Http.DefaultRequestHeaders.Authorization = New AuthenticationHeaderValue("Bearer", token)
                End If

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

                    If root.TryGetProperty("published_at", Nothing) Then
                        info.PublishedAt = ParseUtcStamp(root.GetProperty("published_at").GetString())
                    End If

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

                    ' Same-version republish detection. Every publish stamps the build with a UTC
                    ' timestamp (InformationalVersion "2.1.6+<stamp>") and appends the same stamp
                    ' to the release body. A release whose stamp is newer than this build's own
                    ' stamp is a republish under the same version, and the update must be offered.
                    Dim localBuildStamp = GetLocalBuildStamp()
                    Dim releaseBuildStamp = ParseBuildStamp(info.ReleaseNotes)

                    ' Locate our executable asset in the release by name: the first .exe on a
                    ' release must be the app, or a stranger's binary gets installed wholesale.
                    Dim downloadUrl As String = ""
                    If root.TryGetProperty("assets", Nothing) Then
                        For Each asset In root.GetProperty("assets").EnumerateArray()
                            Dim name = asset.GetProperty("name").GetString()
                            If name IsNot Nothing AndAlso name.StartsWith("BTA_OSG", StringComparison.OrdinalIgnoreCase) AndAlso name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase) Then
                                downloadUrl = asset.GetProperty("browser_download_url").GetString()
                                Exit For
                            End If
                        Next
                    End If
                    info.DownloadUrl = downloadUrl

                    info.Success = True
                    Dim versionDelta = (latestVer IsNot Nothing AndAlso latestVer > currentVer)
                    Dim sameVersion = IsSameVersion(latestVer, currentVer)
                    Dim republish = (sameVersion AndAlso IsRepublishNewer(releaseBuildStamp, localBuildStamp))
                    info.HasUpdate = (versionDelta OrElse republish) AndAlso Not String.IsNullOrEmpty(downloadUrl)
                    Return info
                End Using
            Catch ex As Exception
                info.Success = False
                info.ErrorMessage = ex.Message
                Return info
            End Try
        End Function

        Public Class ReleaseHistoryEntry
        Public Property TagName As String = ""
        Public Property Body As String = ""
        Public Property PublishedAt As DateTime = DateTime.MinValue
    End Class

    ''' <summary>
    ''' Past releases, newest first. The GitHub releases list is the changelog's single
    ''' source of truth; the dialog's local curated history is only the offline fallback
    ''' and the pre-2.1.0 era that predates publishing releases.
    ''' </summary>
    Public Async Function GetRecentReleasesAsync() As Task(Of List(Of ReleaseHistoryEntry))
        Dim entries As New List(Of ReleaseHistoryEntry)()
        Try
            Http.DefaultRequestHeaders.UserAgent.Clear()
            Http.DefaultRequestHeaders.UserAgent.Add(New ProductInfoHeaderValue("BTA_OSG_Updater", "1.0"))

            Dim response = Await Http.GetAsync($"https://api.github.com/repos/{GitHubRepo}/releases").ConfigureAwait(False)
            If Not response.IsSuccessStatusCode Then Return entries

            Dim json = Await response.Content.ReadAsStringAsync().ConfigureAwait(False)
            Using doc = JsonDocument.Parse(json)
                For Each rel In doc.RootElement.EnumerateArray()
                    Dim tag = ""
                    If rel.TryGetProperty("tag_name", Nothing) Then tag = rel.GetProperty("tag_name").GetString()
                    If String.IsNullOrWhiteSpace(tag) Then Continue For

                    Dim body = ""
                    If rel.TryGetProperty("body", Nothing) AndAlso rel.GetProperty("body").ValueKind = JsonValueKind.String Then
                        body = rel.GetProperty("body").GetString()
                    End If
                    Dim pubAt As DateTime = DateTime.MinValue
                    If rel.TryGetProperty("published_at", Nothing) Then
                        pubAt = ParseUtcStamp(rel.GetProperty("published_at").GetString())
                    End If
                    entries.Add(New ReleaseHistoryEntry With {.TagName = tag, .Body = If(body, ""), .PublishedAt = pubAt})
                Next
            End Using
        Catch ex As Exception
            ' Offline or rate-limited: an empty list simply tells the dialog to fall back
            ' to its local curated history, which covers every version including 2.1.4.
        End Try
        Return entries
    End Function

    ''' <summary>
    ''' True when a release tag and an installed assembly version describe the same release
    ''' despite different part counts ("2.2" vs "2.2.0.0"); missing parts read as zero.
    ''' </summary>
    Friend Function IsSameVersion(releaseVer As Version, currentVer As Version) As Boolean
        If releaseVer Is Nothing OrElse currentVer Is Nothing Then Return False
        Return releaseVer.Major = currentVer.Major AndAlso
               Math.Max(releaseVer.Minor, 0) = Math.Max(currentVer.Minor, 0) AndAlso
               Math.Max(releaseVer.Build, 0) = Math.Max(currentVer.Build, 0) AndAlso
               Math.Max(releaseVer.Revision, 0) = Math.Max(currentVer.Revision, 0)
    End Function

    ''' <summary>
    ''' Version as the office reads it: always two parts, so a 2.1.6.0 build reads "2.1"
    ''' while a two-part release reads "2.2". One scheme everywhere versions are shown.
    ''' </summary>
    Friend Function FormatVersion(v As Version) As String
        If v Is Nothing Then Return ""
        Return v.Major & "." & v.Minor
    End Function

    ''' <summary>
    ''' Release tag as the office reads it: "v2.1.5" and "v2.1.6" both read "v2.1".
    ''' Unparsable tags pass through untouched.
    ''' </summary>
    Friend Function FormatTagLabel(tagName As String) As String
        Dim parsed As Version = Nothing
        If Version.TryParse(If(tagName, "").Trim().TrimStart("v"c, "V"c), parsed) Then
            Return "v" & FormatVersion(parsed)
        End If
        Return If(tagName, "")
    End Function

    ''' <summary>
    ''' True when a same-version release stamped releaseStamp must still be treated as an
    ''' update over an installed build carrying localStamp (the release is a republish).
    ''' Either stamp missing keeps the plain version comparison.
    ''' </summary>
    Friend Function IsRepublishNewer(releaseStamp As DateTime, localStamp As DateTime) As Boolean
        Return releaseStamp <> DateTime.MinValue AndAlso localStamp <> DateTime.MinValue AndAlso releaseStamp > localStamp
    End Function

    ''' <summary>
    ''' Parse a timestamp string ("2026-10-03T14:22:11Z" or GitHub's ISO form) as UTC,
    ''' or MinValue when absent or unparsable.
    ''' </summary>
    Friend Function ParseUtcStamp(text As String) As DateTime
        Dim stamp As DateTime
        If DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, stamp) Then Return stamp
        Return DateTime.MinValue
    End Function

    ''' <summary>
    ''' The stamp inside an InformationalVersion string, the part after '+' up to the SDK's
    ''' "." source-revision suffix ("2.1.6+2026-10-03T14:22:11Z.d196c72"). MinValue when the
    ''' build carries no stamp.
    ''' </summary>
    Friend Function ParseInformationalStamp(infoVer As String) As DateTime
        If String.IsNullOrEmpty(infoVer) Then Return DateTime.MinValue
        Dim plus = infoVer.IndexOf("+"c)
        If plus < 0 Then Return DateTime.MinValue

        Dim suffix = infoVer.Substring(plus + 1)
        Dim dot = suffix.IndexOf("."c)
        If dot >= 0 Then suffix = suffix.Substring(0, dot)
        Return ParseUtcStamp(suffix)
    End Function

    ''' <summary>
    ''' The stamp of the running build: InformationalVersion build metadata when present
    ''' ("2.1.6+2026-10-03T14:22:11Z"), else the exe file's write time so copies installed
    ''' before stamping existed can still detect a republish.
    ''' </summary>
    Friend Function GetLocalBuildStamp() As DateTime
        Dim attr = Assembly.GetExecutingAssembly().GetCustomAttribute(Of AssemblyInformationalVersionAttribute)()
        Dim stamp = ParseInformationalStamp(If(attr IsNot Nothing, attr.InformationalVersion, ""))
        If stamp <> DateTime.MinValue Then Return stamp

        Try
            Dim exePath = Environment.ProcessPath
            If Not String.IsNullOrEmpty(exePath) AndAlso File.Exists(exePath) Then
                Return File.GetLastWriteTimeUtc(exePath)
            End If
        Catch
            ' Best-effort fallback; without a stamp the same-version republish check is skipped.
        End Try
        Return DateTime.MinValue
    End Function

    ''' <summary>
    ''' The build stamp publish-release.ps1 writes into the release body as an HTML comment,
    ''' or MinValue for releases that predate stamping.
    ''' </summary>
    Friend Function ParseBuildStamp(body As String) As DateTime
        Const marker As String = "<!-- build:"
        If String.IsNullOrEmpty(body) Then Return DateTime.MinValue

        Dim start = body.IndexOf(marker, StringComparison.Ordinal)
        If start < 0 Then Return DateTime.MinValue

        Dim finish = body.IndexOf("-->", start + marker.Length, StringComparison.Ordinal)
        If finish < 0 Then Return DateTime.MinValue

        Return ParseUtcStamp(body.Substring(start + marker.Length, finish - start - marker.Length).Trim())
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
                Dim safeCurrentExe = currentExe.Replace("'", "''")

                Dim script = "@echo off" & vbCrLf &
                             ":wait" & vbCrLf &
                             $"tasklist /fi ""PID eq {currentPid}"" | findstr ""{currentPid}"" >nul" & vbCrLf &
                             "if not errorlevel 1 (timeout /t 1 /nobreak >nul & goto wait)" & vbCrLf &
                             "timeout /t 1 /nobreak >nul" & vbCrLf &
                             ":copyloop" & vbCrLf &
                             $"copy /y ""{tempExe}"" ""{currentExe}"" >nul" & vbCrLf &
                             "if errorlevel 1 (" & vbCrLf &
                             "    timeout /t 1 /nobreak >nul" & vbCrLf &
                             "    goto copyloop" & vbCrLf &
                             ")" & vbCrLf &
                             $"powershell -NoProfile -Command ""Unblock-File '{safeCurrentExe}'"" >nul 2>&1" & vbCrLf &
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
                Environment.Exit(0)
            Catch ex As Exception
                MessageBox.Show(ownerForm, "Failed to download and apply update: " & ex.Message, "Update Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Function

        Public Async Function CheckAndApplyUpdateAsync(ownerForm As Form, manualCheck As Boolean, Optional onChecked As Action = Nothing) As Task
            Dim info = Await CheckUpdateInfoAsync().ConfigureAwait(True)
            If onChecked IsNot Nothing Then onChecked()
            Using dlg As New FormWhatsNew(info.HasUpdate, info.LatestVersion, info.DownloadUrl, info.ReleaseNotes, info.PublishedAt, alreadyChecked:=True, checkSucceeded:=info.Success)
                dlg.ShowDialog(ownerForm)
            End Using
        End Function
    End Module
End Namespace
