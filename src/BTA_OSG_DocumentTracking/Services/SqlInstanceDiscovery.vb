Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Threading.Tasks
Imports Microsoft.Data.SqlClient
Imports Microsoft.Win32

Namespace BTA_OSG
    ''' <summary>
    ''' Network assist for the connect dialog: lists the SQL Servers a workstation can
    ''' see, then keeps only the ones that actually host BTA_OSG_DB. Discovery never
    ''' saves and never configures silently; a found server is a suggestion the
    ''' operator still confirms by connecting.
    ''' </summary>
    Public Class SqlInstanceDiscovery
        Public Enum ProbeResult As Integer
            FoundDatabase
            LoginRejected
            NoDatabase
            Unreachable
        End Enum

        Public Class Candidate
            Public Property Server As String
            Public Property Port As Integer?
        End Class

        Public Class ScanResult
            Public Property Server As String
            Public Property Port As Integer?
            Public Property Result As ProbeResult
        End Class

        ''' <summary>
        ''' Asks the SQL Server Browser service (UDP 1434) which instances exist on the
        ''' subnet. The enumerator is genuinely flaky behind firewalls, so its answer is an
        ''' assist only: the local machine is always a candidate, because at this scale the
        ''' server is often one of the seats, and an empty list simply means "type it".
        ''' </summary>
        Public Shared Function EnumerateCandidates() As List(Of Candidate)
            Dim list As New List(Of Candidate)()
            Try
                Dim table = Microsoft.Data.Sql.SqlDataSourceEnumerator.Instance.GetDataSources()
                For Each row As DataRow In table.Rows
                    Dim host = row("ServerName").ToString()
                    Dim inst = If(table.Columns.Contains("InstanceName"), row("InstanceName").ToString(), "")
                    Dim port As Integer? = Nothing
                    If table.Columns.Contains("Port") AndAlso Not IsDBNull(row("Port")) Then
                        Dim parsed As Integer
                        If Integer.TryParse(row("Port").ToString(), parsed) Then port = parsed
                    End If
                    Dim serverText = host
                    If Not port.HasValue AndAlso inst.Length > 0 AndAlso Not inst.Equals("MSSQLSERVER", StringComparison.OrdinalIgnoreCase) Then
                        serverText = host & "\" & inst
                    End If
                    list.Add(New Candidate With {.Server = serverText, .Port = port})
                Next
            Catch ex As Exception
                System.Diagnostics.Trace.TraceWarning("SQL instance enumeration failed: " & ex.Message)
            End Try

            Dim local = Environment.MachineName
            ' The local machine is always a candidate, because at this scale the server is
            ' often one of the seats. Its installed instance names come from the registry:
            ' a fresh Express install answers over shared memory even with TCP and the
            ' Browser service still off, and only the exact named instance reaches it.
            For Each inst In LocalInstanceNames()
                Dim target = If(inst.Equals("MSSQLSERVER", StringComparison.OrdinalIgnoreCase), local, local & "\" & inst)
                Dim alreadyListed As Boolean = False
                For Each c In list
                    If c.Server.Equals(target, StringComparison.OrdinalIgnoreCase) Then
                        alreadyListed = True
                        Exit For
                    End If
                Next
                If Not alreadyListed Then list.Add(New Candidate With {.Server = target, .Port = Nothing})
            Next

            Dim hasLocal As Boolean = False
            For Each c In list
                If c.Server.Equals(local, StringComparison.OrdinalIgnoreCase) OrElse c.Server.StartsWith(local & "\", StringComparison.OrdinalIgnoreCase) Then
                    hasLocal = True
                    Exit For
                End If
            Next
            If Not hasLocal Then list.Add(New Candidate With {.Server = local, .Port = Nothing})
            Return list
        End Function

        ''' <summary>
        ''' Friendly names of the SQL Server instances installed on this machine, read
        ''' from the Instance Names registry key. The 64-bit view is read explicitly
        ''' because a 32-bit process would otherwise see the WOW64 view, where the key
        ''' does not exist.
        ''' </summary>
        Public Shared Function LocalInstanceNames() As List(Of String)
            Dim names As New List(Of String)()
            Dim views = If(Environment.Is64BitProcess,
                           New RegistryView() {RegistryView.Default},
                           New RegistryView() {RegistryView.Default, RegistryView.Registry64})
            For Each view In views
                Try
                    Using base As RegistryKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view)
                        Using key = base.OpenSubKey("SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL")
                            If key Is Nothing Then Continue For
                            For Each name As String In key.GetValueNames()
                                If Not names.Contains(name) Then names.Add(name)
                            Next
                            Return names
                        End Using
                    End Using
                Catch ex As Exception
                    ' An unreadable key means the answer is "none"; the caller only uses
                    ' this to probe local candidates and to gate the prepare button.
                End Try
            Next
            Return names
        End Function

        ''' <summary>
        ''' True when this machine hosts any SQL Server instance. Gates the connect
        ''' dialog's "Set up this server" offer: provisioning can only run on a machine
        ''' that actually has SQL Server to administer.
        ''' </summary>
        Public Shared Function HasLocalSqlServer() As Boolean
            Return LocalInstanceNames().Count > 0
        End Function

        ''' <summary>
        ''' One connection attempt against BTA_OSG_DB with the wizard's SQL login. The three
        ''' error classes an operator actually meets are told apart: the database is there
        ''' and complete (sentinel tables), the server exists but this login is not granted
        ''' yet (18456), and the server has no BTA_OSG_DB at all (4060).
        ''' </summary>
        Public Shared Function Probe(server As String, port As Integer?, user As String, password As String) As ProbeResult
            Dim b As New SqlConnectionStringBuilder()
            b.DataSource = If(port.HasValue, server & "," & port.Value.ToString(), server)
            b.InitialCatalog = "BTA_OSG_DB"
            b.ConnectTimeout = 2
            b.Encrypt = True
            b.TrustServerCertificate = True
            b.IntegratedSecurity = False
            b.UserID = user
            b.Password = password

            Try
                Using conn As New SqlConnection(b.ConnectionString)
                    conn.Open()
                    Using cmd As New SqlCommand("SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME IN ('tbl_Documents','tbl_Users','tbl_AuditTrail')", conn)
                        Dim sentinelCount = Convert.ToInt32(cmd.ExecuteScalar())
                        Return If(sentinelCount >= 3, ProbeResult.FoundDatabase, ProbeResult.NoDatabase)
                    End Using
                End Using
            Catch ex As SqlException
                For i As Integer = 0 To ex.Errors.Count - 1
                    If ex.Errors(i).Number = 4060 Then Return ProbeResult.NoDatabase
                    If ex.Errors(i).Number = 18456 Then Return ProbeResult.LoginRejected
                Next
                Return ProbeResult.Unreachable
            Catch ex As Exception
                System.Diagnostics.Trace.TraceWarning("SQL probe of " & server & " failed: " & ex.Message)
                Return ProbeResult.Unreachable
            End Try
        End Function

        ''' <summary>
        ''' Probes every candidate in parallel. A dead host costs its 2s connect timeout, so
        ''' the whole scan is capped at 10s: a large silent subnet must not freeze the wizard.
        ''' </summary>
        Public Shared Function Discover(user As String, password As String) As List(Of ScanResult)
            Dim candidates = EnumerateCandidates()
            Dim tasks As New List(Of Task(Of ScanResult))()
            For Each cand In candidates
                tasks.Add(Task.Run(Function() New ScanResult With {
                                       .Server = cand.Server,
                                       .Port = cand.Port,
                                       .Result = Probe(cand.Server, cand.Port, user, password)
                                   }))
            Next
            Task.WaitAll(tasks.ToArray(), TimeSpan.FromSeconds(10))

            Dim results As New List(Of ScanResult)()
            For Each t In tasks
                If t.IsCompleted AndAlso Not t.IsFaulted AndAlso Not t.IsCanceled AndAlso t.Result IsNot Nothing Then
                    results.Add(t.Result)
                End If
            Next
            Return results
        End Function
    End Class
End Namespace
