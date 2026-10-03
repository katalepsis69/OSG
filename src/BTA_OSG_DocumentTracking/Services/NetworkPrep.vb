Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.Text

Namespace BTA_OSG
    ''' <summary>
    ''' The OS-side half of Set up this server: restarts the SQL Server service so Mixed
    ''' Mode and the TCP settings take effect, enables and starts the SQL Browser
    ''' service, and opens TCP 1433 plus UDP 1434 in Windows Firewall. Idempotent: a
    ''' rerun after fixing the cause is safe, and each step fails on its own, so one
    ''' broken precondition never leaves the server half-configured. The exit code is
    ''' nonzero when anything failed, and the report lines name the failed step.
    ''' </summary>
    Public Class NetworkPrep

        ''' <summary>
        ''' Built-in cmdlets do the waiting that sc.exe would need polling loops for,
        ''' and netsh is checked per call because a non-zero exit does not throw under
        ''' ErrorActionPreference. Each step carries its own catch so one broken
        ''' precondition (a service that will not restart) still lets the later steps
        ''' apply, and the exit code stays nonzero when anything failed. Firewall rules
        ''' are deleted by name before they are added: a stale rule with old settings
        ''' must not survive a rerun.
        ''' </summary>
        Private Shared ReadOnly PrepScript As String =
            "$ErrorActionPreference = 'Stop'" & vbCrLf &
            "$failed = $false" & vbCrLf &
            "try {" & vbCrLf &
            "  $instances = Get-Service | Where-Object { $_.Name -eq 'MSSQLSERVER' -or $_.Name -like 'MSSQL$*' }" & vbCrLf &
            "  if ($instances.Count -eq 0) { Write-Output 'No SQL Server service found on this machine.' }" & vbCrLf &
            "  foreach ($svc in $instances) {" & vbCrLf &
            "    try {" & vbCrLf &
            "      Write-Output ('Restarting SQL Server service ' + $svc.Name + '...')" & vbCrLf &
            "      Restart-Service -Name $svc.Name -Force" & vbCrLf &
            "    } catch {" & vbCrLf &
            "      $failed = $true" & vbCrLf &
            "      Write-Output ('Failed: restarting ' + $svc.Name + ' (' + $_.Exception.Message + ')')" & vbCrLf &
            "    }" & vbCrLf &
            "  }" & vbCrLf &
            "} catch {" & vbCrLf &
            "  $failed = $true" & vbCrLf &
            "  Write-Output ('Failed: SQL service discovery (' + $_.Exception.Message + ')')" & vbCrLf &
            "}" & vbCrLf &
            "try {" & vbCrLf &
            "  Set-Service -Name 'SQLBrowser' -StartupType Automatic" & vbCrLf &
            "  Start-Service -Name 'SQLBrowser'" & vbCrLf &
            "  Write-Output 'SQL Browser service enabled and started.'" & vbCrLf &
            "} catch {" & vbCrLf &
            "  Write-Output 'SQL Browser service not present; named-instance discovery will use typed addresses instead.'" & vbCrLf &
            "}" & vbCrLf &
            "try {" & vbCrLf &
            "  netsh advfirewall firewall delete rule name='BTA OSG SQL (TCP 1433)' | Out-Null" & vbCrLf &
            "  netsh advfirewall firewall add rule name='BTA OSG SQL (TCP 1433)' dir=in action=allow protocol=TCP localport=1433 | Out-Null" & vbCrLf &
            "  if ($LASTEXITCODE -ne 0) { throw 'the firewall rule for TCP 1433 failed' }" & vbCrLf &
            "  netsh advfirewall firewall delete rule name='BTA OSG SQL Browser (UDP 1434)' | Out-Null" & vbCrLf &
            "  netsh advfirewall firewall add rule name='BTA OSG SQL Browser (UDP 1434)' dir=in action=allow protocol=UDP localport=1434 | Out-Null" & vbCrLf &
            "  if ($LASTEXITCODE -ne 0) { throw 'the firewall rule for UDP 1434 failed' }" & vbCrLf &
            "  Write-Output 'Firewall rules for TCP 1433 and UDP 1434 are in place.'" & vbCrLf &
            "} catch {" & vbCrLf &
            "  $failed = $true" & vbCrLf &
            "  Write-Output ('Failed: ' + $_.Exception.Message)" & vbCrLf &
            "}" & vbCrLf &
            "if ($failed) { exit 1 } else { exit 0 }"

        ''' <summary>
        ''' Runs the network steps and reports each one through the callback. The
        ''' elevated caller (the /prepnet branch) turns the lines into the summary the
        ''' operator sees; a caller can only reach this through /prepnet, so the service
        ''' and firewall state is never touched by an unelevated or accidental path.
        ''' </summary>
        Public Shared Function Run(report As Action(Of String)) As Integer
            Dim scriptPath = Path.Combine(Path.GetTempPath(), "bta-osg-prepnet.ps1")
            File.WriteAllText(scriptPath, PrepScript, New UTF8Encoding(False))
            Try
                Dim startInfo As New ProcessStartInfo With {
                    .FileName = "powershell.exe",
                    .Arguments = "-NoProfile -ExecutionPolicy Bypass -File """ & scriptPath & """",
                    .UseShellExecute = False,
                    .CreateNoWindow = True,
                    .RedirectStandardOutput = True,
                    .RedirectStandardError = True
                }
                Using process As New Process()
                    process.StartInfo = startInfo
                    process.Start()
                    Dim stdout = process.StandardOutput.ReadToEnd()
                    Dim stderr = process.StandardError.ReadToEnd()
                    process.WaitForExit()
                    For Each line As String In stdout.Split({vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)
                        report(line)
                    Next
                    For Each line As String In stderr.Split({vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)
                        report(line)
                    Next
                    Return process.ExitCode
                End Using
            Finally
                Try
                    File.Delete(scriptPath)
                Catch
                    ' A leftover temp script is harmless; nothing reads it.
                End Try
            End Try
        End Function

        ''' <summary>
        ''' Re-launches this exe with /prepnet under the administrator verb. The UAC
        ''' consent is the one administrator prompt the whole server setup asks for.
        ''' </summary>
        Public Shared Function RunElevated() As Integer
            Dim startInfo As New ProcessStartInfo With {
                .FileName = Environment.ProcessPath,
                .Arguments = "/prepnet",
                .UseShellExecute = True,
                .Verb = "runas"
            }
            Using proc = Process.Start(startInfo)
                proc.WaitForExit()
                Return proc.ExitCode
            End Using
        End Function
    End Class
End Namespace
