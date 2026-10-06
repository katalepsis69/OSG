Option Explicit On
Option Strict On

Imports System
Imports System.Diagnostics
Imports System.IO
Imports System.Net.Http
Imports System.Net.Sockets
Imports System.Threading.Tasks
Imports System.Windows.Forms

Namespace BTA_OSG
    ''' <summary>
    ''' Manages on-demand background execution of the local OSGPortal PHP server
    ''' using Laragon's PHP engine for this project, and launches the browser when requested.
    ''' </summary>
    Public Class PortalServerManager
        Private Shared _phpProcess As Process = Nothing
        Private Shared _cloudflaredProcess As Process = Nothing
        Private Shared _mysqlProcess As Process = Nothing
        Private Shared _activeTunnelUrl As String = Nothing
        Private Shared ReadOnly _syncLock As New Object()
        Private Shared _isHooked As Boolean = False

        ''' <summary>
        ''' Ensures the OSGPortal server (bundled MySQL + PHP) is running headlessly in the background.
        ''' Does not open the web browser.
        ''' </summary>
        Public Shared Async Function EnsureRunningAsync(Optional statusCallback As Action(Of String) = Nothing) As Task(Of Boolean)
            Dim baseUrl = AppSettings.Instance.PortalSettings.BaseUrl
            ' The desktop portal is a local service (the headless server binds 127.0.0.1): a
            ' settings layer that still carries an external default must not make the badge
            ' probe a stranger's server, so a non-loopback BaseUrl is replaced outright.
            If String.IsNullOrWhiteSpace(baseUrl) OrElse Not IsLoopbackBaseUrl(baseUrl) Then
                baseUrl = "http://localhost:8085"
            End If

            ' 1. Check if already responding
            Dim isRunning = Await CheckServerRespondingAsync(baseUrl).ConfigureAwait(False)
            If isRunning Then
                statusCallback?.Invoke($"OSGPortal: Active ({baseUrl})")
                Return True
            End If

            statusCallback?.Invoke("OSGPortal: Auto-starting local service...")
            EnsureMySqlRunning()
            Dim started = StartHeadlessServer()
            If started Then
                For i As Integer = 1 To 6
                    Await Task.Delay(500).ConfigureAwait(False)
                    isRunning = Await CheckServerRespondingAsync(baseUrl).ConfigureAwait(False)
                    If isRunning Then
                        statusCallback?.Invoke($"OSGPortal: Active ({baseUrl})")
                        Return True
                    End If
                Next
            End If

            statusCallback?.Invoke("OSGPortal: Standby (Offline)")
            Return False
        End Function

        ''' <summary>
        ''' Ensures the OSGPortal server is running (starting it headlessly if needed),
        ''' then opens the portal URL in the default web browser.
        ''' </summary>
        Public Shared Async Function EnsureRunningAndOpenAsync(Optional statusCallback As Action(Of String) = Nothing) As Task(Of Boolean)
            Dim baseUrl = AppSettings.Instance.PortalSettings.BaseUrl
            If String.IsNullOrWhiteSpace(baseUrl) Then
                baseUrl = "http://localhost:8085"
            End If

            statusCallback?.Invoke("OSGPortal: Checking server status...")
            Dim isRunning = Await EnsureRunningAsync(statusCallback).ConfigureAwait(False)

            ' Open in user's default browser
            Try
                Process.Start(New ProcessStartInfo With {
                    .FileName = baseUrl,
                    .UseShellExecute = True
                })
                Return True
            Catch ex As Exception
                statusCallback?.Invoke("OSGPortal: Failed to open browser: " & ex.Message)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' True only for loopback hosts: the local portal posture never points at a
        ''' remote machine, so anything else is a stale or foreign setting.
        ''' </summary>
        Private Shared Function IsLoopbackBaseUrl(baseUrl As String) As Boolean
            Try
                Dim host = New Uri(baseUrl).Host.ToLowerInvariant()
                Return host = "localhost" OrElse host = "127.0.0.1" OrElse host = "::1"
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Pings the base URL to see if it responds.
        ''' </summary>
        Public Shared Async Function CheckServerRespondingAsync(url As String) As Task(Of Boolean)
            Try
                Using client As New HttpClient() With {.Timeout = TimeSpan.FromSeconds(1)}
                    client.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent", "BTA-OSG-PortalBridge/2.1")
                    Using response = Await client.GetAsync(url, HttpCompletionOption.ResponseHeadersRead).ConfigureAwait(False)
                        Return True
                    End Using
                End Using
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Starts PHP built-in server in the background using Laragon's PHP binary for the portal folder only.
        ''' </summary>
        Public Shared Function StartHeadlessServer() As Boolean
            SyncLock _syncLock
                If _phpProcess IsNot Nothing AndAlso Not _phpProcess.HasExited Then
                    Return True
                End If

                Dim publicDir = ResolvePortalPublicDir()
                If String.IsNullOrEmpty(publicDir) OrElse Not Directory.Exists(publicDir) Then
                    Return False
                End If

                Dim phpExe = ResolvePhpExecutable()
                If String.IsNullOrEmpty(phpExe) OrElse Not File.Exists(phpExe) Then
                    Return False
                End If

                ' Bind to loopback only: internet exposure goes through the cloudflared tunnel
                ' (which connects from 127.0.0.1), and a 0.0.0.0 bind would put the portal on
                ' every LAN interface of the workstation.
                Dim host = "127.0.0.1:8085"
                Try
                    Dim uri As New Uri(AppSettings.Instance.PortalSettings.BaseUrl)
                    If uri.Port > 0 Then
                        host = "127.0.0.1:" & uri.Port
                    End If
                Catch
                    host = "127.0.0.1:8085"
                End Try

                Dim psi As New ProcessStartInfo With {
                    .FileName = phpExe,
                    .Arguments = BuildPhpArguments(phpExe, host, publicDir),
                    .CreateNoWindow = True,
                    .UseShellExecute = False,
                    .WindowStyle = ProcessWindowStyle.Hidden,
                    .WorkingDirectory = publicDir
                }

                ' Set environment variables for database and bridge credentials
                psi.Environment("PORTAL_DB_HOST") = "127.0.0.1"
                psi.Environment("PORTAL_DB_PORT") = "3306"
                psi.Environment("PORTAL_DB_NAME") = "osg"
                psi.Environment("PORTAL_DB_USER") = "root"
                psi.Environment("PORTAL_DB_PASS") = ""
                psi.Environment("BRIDGE_SECRET_KEY") = AppSettings.Instance.PortalSettings.BridgeKey
                psi.Environment("PORTAL_BASE_URL") = AppSettings.Instance.PortalSettings.BaseUrl
                ' Mail credentials come from the encrypted settings file (Portal.BrevoApiKey
                ' and friends). No live key ships in source; an empty key simply means the
                ' portal cannot send mail until the operator configures one.
                psi.Environment("BREVO_API_KEY") = AppSettings.Instance.PortalSettings.BrevoApiKey
                psi.Environment("BREVO_SENDER_EMAIL") = AppSettings.Instance.PortalSettings.BrevoSenderEmail
                psi.Environment("BREVO_SENDER_NAME") = AppSettings.Instance.PortalSettings.BrevoSenderName

                Try
                    _phpProcess = Process.Start(psi)
                    If Not _isHooked Then
                        _isHooked = True
                        AddHandler AppDomain.CurrentDomain.ProcessExit, AddressOf OnProcessExit
                        AddHandler Application.ApplicationExit, AddressOf OnProcessExit
                    End If
                    Return True
                Catch ex As Exception
                    Trace.WriteLine("Failed to spawn headless OSGPortal PHP: " & ex.Message)
                    Return False
                End Try
            End SyncLock
        End Function

        ''' <summary>
        ''' The bundled runtime ships a minimal php.ini (expose_php off, display_errors off);
        ''' point the built-in server at it when present so compiled defaults never leak
        ''' version banners or error output.
        ''' </summary>
        Private Shared Function BuildPhpArguments(phpExe As String, host As String, publicDir As String) As String
            Dim args = ""
            Dim iniPath = Path.Combine(Path.GetDirectoryName(phpExe), "php.ini")
            If File.Exists(iniPath) Then
                args &= $"-c ""{iniPath}"" "
            End If
            args &= $"-S {host} -t ""{publicDir}"""
            Return args
        End Function

        ''' <summary>
        ''' Stops the spawned background PHP server if running.
        ''' </summary>
        Public Shared Sub StopServer()
            SyncLock _syncLock
                If _phpProcess IsNot Nothing Then
                    Try
                        If Not _phpProcess.HasExited Then
                            _phpProcess.Kill(True)
                        End If
                    Catch
                    Finally
                        _phpProcess.Dispose()
                        _phpProcess = Nothing
                    End Try
                End If

                If _mysqlProcess IsNot Nothing Then
                    ' The bundled MySQL belongs to this app's portal data directory, so it
                    ' is ours to stop; a MySQL this process did not start is left alone.
                    Try
                        If Not _mysqlProcess.HasExited Then
                            _mysqlProcess.Kill(True)
                        End If
                    Catch
                    Finally
                        _mysqlProcess.Dispose()
                        _mysqlProcess = Nothing
                    End Try
                End If

                If _cloudflaredProcess IsNot Nothing Then
                    Try
                        If Not _cloudflaredProcess.HasExited Then
                            _cloudflaredProcess.Kill(True)
                        End If
                    Catch
                    Finally
                        _cloudflaredProcess.Dispose()
                        _cloudflaredProcess = Nothing
                        _activeTunnelUrl = Nothing
                    End Try
                End If
            End SyncLock
        End Sub

        ''' <summary>
        ''' Gets the active trycloudflare.com tunnel URL if running.
        ''' </summary>
        Public Shared ReadOnly Property ActiveTunnelUrl As String
            Get
                Return _activeTunnelUrl
            End Get
        End Property

        ''' <summary>
        ''' Gets whether the Cloudflare tunnel process is currently running.
        ''' </summary>
        Public Shared ReadOnly Property IsTunnelRunning As Boolean
            Get
                SyncLock _syncLock
                    Return _cloudflaredProcess IsNot Nothing AndAlso Not _cloudflaredProcess.HasExited
                End SyncLock
            End Get
        End Property

        ''' <summary>
        ''' Stops the running Cloudflare tunnel process if active.
        ''' </summary>
        Public Shared Sub StopTunnel()
            SyncLock _syncLock
                If _cloudflaredProcess IsNot Nothing Then
                    Try
                        If Not _cloudflaredProcess.HasExited Then
                            _cloudflaredProcess.Kill(True)
                        End If
                    Catch
                    Finally
                        _cloudflaredProcess.Dispose()
                        _cloudflaredProcess = Nothing
                        _activeTunnelUrl = Nothing
                    End Try
                End If
            End SyncLock
        End Sub

        Private Shared Sub OnProcessExit(sender As Object, e As EventArgs)
            StopServer()
        End Sub

        Private Shared Sub CaptureTunnelUrlFromOutput(sender As Object, e As DataReceivedEventArgs)
            If String.IsNullOrEmpty(e.Data) Then Return
            Dim m = System.Text.RegularExpressions.Regex.Match(e.Data, "https://[a-zA-Z0-9-]+\.trycloudflare\.com")
            If m.Success Then
                _activeTunnelUrl = m.Value
            End If
        End Sub

        ' The desktop can run from the repo's bin tree, a published dist folder, or beside the
        ' portal checkout; probe the same relative locations for each bundled artifact.
        Private Shared Function RuntimePathCandidates(ParamArray relativeParts As String()) As String()
            Dim relative = Path.Combine(relativeParts)
            Return New String() {
                Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "..", relative)),
                Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", relative)),
                Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, relative))
            }
        End Function

        Private Shared Sub EnsureMySqlRunning()
            If IsPortListening(3306) Then Return
            If TryStartBundledMySql() Then Return
            TryStartLaragonMySql()
        End Sub

        Private Shared Function TryStartBundledMySql() As Boolean
            For Each baseDir As String In RuntimePathCandidates("portal", "runtime")
                Dim mysqldExe = Path.Combine(baseDir, "mysql", "mysqld.exe")
                If Not File.Exists(mysqldExe) Then
                    mysqldExe = Path.Combine(baseDir, "mysql", "bin", "mysqld.exe")
                End If
                Dim myIni = Path.Combine(baseDir, "mysql", "my.ini")
                Dim dataDir = Path.Combine(baseDir, "data")
                Dim mysqlBase = Path.Combine(baseDir, "mysql")
                Dim shareDir = Path.Combine(baseDir, "mysql", "share")

                If File.Exists(mysqldExe) AndAlso Directory.Exists(dataDir) Then
                    Try
                        Dim args = $"--datadir=""{dataDir}"""
                        If File.Exists(myIni) Then
                            args = $"--defaults-file=""{myIni}"" " & args
                        End If
                        If Directory.Exists(mysqlBase) Then
                            args &= $" --basedir=""{mysqlBase}"""
                        End If
                        If Directory.Exists(shareDir) Then
                            args &= $" --lc-messages-dir=""{shareDir}"""
                        End If
                        _mysqlProcess = Process.Start(New ProcessStartInfo With {
                            .FileName = mysqldExe,
                            .Arguments = args,
                            .CreateNoWindow = True,
                            .UseShellExecute = False,
                            .WindowStyle = ProcessWindowStyle.Hidden
                        })
                        Threading.Thread.Sleep(1500)
                        If IsPortListening(3306) Then Return True
                    Catch
                    End Try
                End If
            Next
            Return False
        End Function

        Private Shared Sub TryStartLaragonMySql()
            Const laragonMysql = "C:\laragon\bin\mysql"
            If Directory.Exists(laragonMysql) Then
                For Each dirPath As String In Directory.GetDirectories(laragonMysql, "mysql-*")
                    Dim cand = Path.Combine(dirPath, "bin", "mysqld.exe")
                    If File.Exists(cand) Then
                        Try
                            ' Tracked like the bundled mysqld so StopServer can kill it; an
                            ' untracked spawn outlives the app and squats on port 3306.
                            _mysqlProcess = Process.Start(New ProcessStartInfo With {
                                .FileName = cand,
                                .CreateNoWindow = True,
                                .UseShellExecute = False,
                                .WindowStyle = ProcessWindowStyle.Hidden
                            })
                            Threading.Thread.Sleep(800)
                            Return
                        Catch
                        End Try
                    End If
                Next
            End If
        End Sub

        Private Shared Function IsPortListening(port As Integer) As Boolean
            Try
                Using client As New TcpClient()
                    Dim ar = client.BeginConnect("127.0.0.1", port, Nothing, Nothing)
                    Dim success = ar.AsyncWaitHandle.WaitOne(TimeSpan.FromMilliseconds(500))
                    If Not success Then Return False
                    client.EndConnect(ar)
                    Return True
                End Using
            Catch
                Return False
            End Try
        End Function

        Private Shared Function ResolvePortalPublicDir() As String
            Dim candidates As New List(Of String)(RuntimePathCandidates("portal", "public"))
            candidates.Add("C:\laragon\www\portal\public")

            For Each cand In candidates
                If Directory.Exists(cand) AndAlso File.Exists(Path.Combine(cand, "index.php")) Then
                    Return cand
                End If
            Next

            Return Nothing
        End Function

        Private Shared Function ResolvePhpExecutable() As String
            For Each cand As String In RuntimePathCandidates("portal", "runtime", "php", "php.exe")
                If File.Exists(cand) Then Return cand
            Next

            ' Fallback to Laragon PHP
            Const laragonPhpDir = "C:\laragon\bin\php"
            If Directory.Exists(laragonPhpDir) Then
                For Each dirPath As String In Directory.GetDirectories(laragonPhpDir, "php-*")
                    Dim cand = Path.Combine(dirPath, "php.exe")
                    If File.Exists(cand) Then Return cand
                Next
            End If

            ' 3. Fallback to XAMPP PHP
            Const xamppPhp = "C:\xampp\php\php.exe"
            If File.Exists(xamppPhp) Then Return xamppPhp

            ' 4. Fallback to system PATH
            Dim pathEnv = Environment.GetEnvironmentVariable("PATH")
            If Not String.IsNullOrEmpty(pathEnv) Then
                For Each p In pathEnv.Split(";"c)
                    If Not String.IsNullOrWhiteSpace(p) Then
                        Dim cand = Path.Combine(p.Trim(), "php.exe")
                        If File.Exists(cand) Then Return cand
                    End If
                Next
            End If

            Return Nothing
        End Function

        ''' <summary>
        ''' Locates the bundled cloudflared executable.
        ''' </summary>
        Public Shared Function ResolveCloudflaredExecutable() As String
            For Each fileName In {"cloudflared.exe", "cloudflared-windows-amd64.exe"}
                For Each cand In RuntimePathCandidates("portal", fileName)
                    If File.Exists(cand) Then Return cand
                Next
            Next

            Return Nothing
        End Function

        ''' <summary>
        ''' Starts Cloudflare Tunnel headlessly with ZERO popup windows, captures the public URL,
        ''' and stops cleanly on app exit.
        ''' </summary>
        Public Shared Async Function StartCloudflareTunnelHeadlessAsync(Optional statusCallback As Action(Of String) = Nothing) As Task(Of String)
            Try
                If IsTunnelRunning AndAlso Not String.IsNullOrEmpty(_activeTunnelUrl) Then
                    statusCallback?.Invoke($"Cloudflare Live: {_activeTunnelUrl}")
                    Return _activeTunnelUrl
                End If

                EnsureMySqlRunning()
                StartHeadlessServer()

                Dim exePath = ResolveCloudflaredExecutable()
                If String.IsNullOrEmpty(exePath) OrElse Not File.Exists(exePath) Then
                    statusCallback?.Invoke("OSGPortal: cloudflared.exe not found in portal folder.")
                    Return Nothing
                End If

                SyncLock _syncLock
                    If _cloudflaredProcess Is Nothing OrElse _cloudflaredProcess.HasExited Then
                        Dim uriStr = "http://127.0.0.1:8085"
                        Dim psi As New ProcessStartInfo With {
                            .FileName = exePath,
                            .Arguments = $"tunnel --metrics localhost:20241 --url {uriStr}",
                            .CreateNoWindow = True,
                            .UseShellExecute = False,
                            .WindowStyle = ProcessWindowStyle.Hidden,
                            .RedirectStandardError = True,
                            .RedirectStandardOutput = True,
                            .WorkingDirectory = Path.GetDirectoryName(exePath)
                        }

                        _cloudflaredProcess = New Process With {
                            .StartInfo = psi,
                            .EnableRaisingEvents = True
                        }

                        AddHandler _cloudflaredProcess.ErrorDataReceived, AddressOf CaptureTunnelUrlFromOutput
                        AddHandler _cloudflaredProcess.OutputDataReceived, AddressOf CaptureTunnelUrlFromOutput

                        _cloudflaredProcess.Start()
                        _cloudflaredProcess.BeginErrorReadLine()
                        _cloudflaredProcess.BeginOutputReadLine()
                    End If
                End SyncLock

                statusCallback?.Invoke("Cloudflare: Connecting secure tunnel...")

                ' Query the metrics server on http://127.0.0.1:20241/quicktunnel or wait for output regex
                Using client As New HttpClient() With {.Timeout = TimeSpan.FromSeconds(2)}
                    For i As Integer = 1 To 25
                        If Not String.IsNullOrEmpty(_activeTunnelUrl) Then
                            Exit For
                        End If

                        Await Task.Delay(400).ConfigureAwait(False)

                        Try
                            Dim json = Await client.GetStringAsync("http://127.0.0.1:20241/quicktunnel").ConfigureAwait(False)
                            If Not String.IsNullOrWhiteSpace(json) AndAlso json.Contains("hostname") Then
                                Dim idx = json.IndexOf("""hostname"":""", StringComparison.OrdinalIgnoreCase)
                                If idx >= 0 Then
                                    Dim startIdx = idx + 12
                                    Dim endIdx = json.IndexOf(""""c, startIdx)
                                    If endIdx > startIdx Then
                                        Dim host = json.Substring(startIdx, endIdx - startIdx)
                                        _activeTunnelUrl = "https://" & host
                                        Exit For
                                    End If
                                End If
                            End If
                        Catch
                        End Try
                    Next
                End Using

                If Not String.IsNullOrEmpty(_activeTunnelUrl) Then
                    statusCallback?.Invoke($"Cloudflare Live: {_activeTunnelUrl}")
                    Return _activeTunnelUrl
                Else
                    statusCallback?.Invoke("Cloudflare: Tunnel active (Waiting for URL)")
                    Return Nothing
                End If
            Catch ex As Exception
                statusCallback?.Invoke("Failed to launch Cloudflare Tunnel: " & ex.Message)
                Return Nothing
            End Try
        End Function
    End Class
End Namespace
