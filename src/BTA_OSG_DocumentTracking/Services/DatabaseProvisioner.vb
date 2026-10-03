Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Data
Imports Microsoft.Data.SqlClient

Namespace BTA_OSG
    ''' <summary>
    ''' Applies the canonical schema (db/scripts 001-015, embedded in the exe) and creates
    ''' the shared bta_app SQL login, so the connect dialog can prepare a fresh server
    ''' with no scripts on disk. Every step is idempotent: the scripts carry their own
    ''' IF NOT EXISTS / COL_LENGTH guards, so a re-run repairs a partial schema and is a
    ''' clean no-op on a complete one. Provisioning needs a sysadmin connection, which in
    ''' a workgroup office only exists on the server machine itself.
    ''' </summary>
    Public Class DatabaseProvisioner
        Public Const AppDatabaseName As String = "BTA_OSG_DB"

        Private Const ResourcePrefix As String = "osgsql."
        ' 001 creates the database, so it runs against master; the rest run inside it.
        Private Shared ReadOnly CatalogScripts As New List(Of Tuple(Of String, String)) From {
            Tuple.Create("001_create_database.sql", "master"),
            Tuple.Create("002_create_reference_tables.sql", AppDatabaseName),
            Tuple.Create("003_create_security_tables.sql", AppDatabaseName),
            Tuple.Create("004_create_document_tables.sql", AppDatabaseName),
            Tuple.Create("005_create_audit_tables.sql", AppDatabaseName),
            Tuple.Create("006_create_indexes.sql", AppDatabaseName),
            Tuple.Create("007_create_seed_data.sql", AppDatabaseName),
            Tuple.Create("008_osg_target_migration.sql", AppDatabaseName),
            Tuple.Create("009_portal_external_control_number.sql", AppDatabaseName),
            Tuple.Create("010_osg_workstation_alignment.sql", AppDatabaseName),
            Tuple.Create("011_create_workstation_heartbeat.sql", AppDatabaseName),
            Tuple.Create("012_retire_bootstrap_admin.sql", AppDatabaseName),
            Tuple.Create("013_realign_document_sequences.sql", AppDatabaseName),
            Tuple.Create("014_audit_hash_chain.sql", AppDatabaseName),
            Tuple.Create("015_replay_and_integrity_fixes.sql", AppDatabaseName)
        }

        ''' <summary>
        ''' Splits on exact-match uppercase GO, the whole dialect db/scripts use. A batch
        ''' that is only whitespace is dropped so callers never see trailing air after the
        ''' final GO.
        ''' </summary>
        Public Shared Iterator Function SplitBatches(sql As String) As IEnumerable(Of String)
            Dim current As New System.Text.StringBuilder()
            For Each line As String In sql.Replace(vbLf, vbCr).Split(vbCr)
                If line.Trim().ToUpperInvariant() = "GO" Then
                    If current.Length > 0 AndAlso current.ToString().Trim().Length > 0 Then Yield current.ToString()
                    current.Clear()
                Else
                    current.AppendLine(line)
                End If
            Next
            If current.Length > 0 AndAlso current.ToString().Trim().Length > 0 Then Yield current.ToString()
        End Function

        Public Shared Sub Provision(adminConnectionString As String)
            Provision(adminConnectionString, AppDatabaseName)
        End Sub

        ''' <summary>
        ''' Provisions a named catalog instead of the canonical office database. The test
        ''' harness uses this to build a throwaway copy, so probe documents and staff can
        ''' never accumulate in the office database. 001 is skipped (it names the canonical
        ''' database verbatim; the caller creates the catalog first), and every remaining
        ''' script's BTA_OSG_DB references are re-pointed at the named catalog.
        ''' </summary>
        Public Shared Sub Provision(adminConnectionString As String, databaseName As String)
            For Each entry In CatalogScripts
                If entry.Item2 = "master" Then Continue For
                Dim sql = ReadEmbeddedScript(entry.Item1).Replace(AppDatabaseName, databaseName)
                ExecuteBatches(entry.Item1, sql, databaseName, adminConnectionString)
            Next
        End Sub

        ''' <summary>
        ''' Makes SQL-authentication seats work on the server: flips a Windows-only instance
        ''' to Mixed Mode (effective after the SQL service restarts), then creates the login
        ''' if missing or, when it exists, resets its password to the one typed and clears
        ''' any lockout, so a forgotten or re-typed password self-heals in one Prepare click
        ''' instead of stranding every seat. Its user in BTA_OSG_DB and db_owner membership
        ''' follow. CHECK_EXPIRATION is OFF deliberately: an expired login would lock every
        ''' seat out of the office at once.
        ''' </summary>
        Public Shared Sub EnsureBtaAppLogin(adminConnectionString As String, loginName As String, password As String)
            If String.IsNullOrWhiteSpace(loginName) Then Throw New ArgumentException("Login name is required.", NameOf(loginName))
            If password Is Nothing OrElse password.Length < 8 Then Throw New ArgumentException("The SQL login password must be at least 8 characters.", NameOf(password))

            Dim master As New SqlConnectionStringBuilder(adminConnectionString) With {.InitialCatalog = "master"}
            Using conn As New SqlConnection(master.ConnectionString)
                conn.Open()
                ' A fresh Express install answers Windows-authentication-only, and a login
                ' created on it could never sign in: flip the instance to Mixed Mode. The
                ' change takes effect when the SQL Server service next restarts.
                Using cmd As New SqlCommand(
                    "IF CAST(SERVERPROPERTY('IsIntegratedSecurityOnly') AS int) = 1 " &
                    "EXEC xp_instance_regwrite N'HKEY_LOCAL_MACHINE', N'Software\Microsoft\MSSQLServer\MSSQLServer', N'LoginMode', REG_DWORD, 2", conn)
                    cmd.ExecuteNonQuery()
                End Using
                ' CREATE LOGIN/ALTER LOGIN's grammar accepts a literal password only (a
                ' variable is rejected even through sp_executesql), so the value is
                ' quote-escaped into the dynamic statement instead of bound as a parameter;
                ' every other statement in this codebase stays fully parameterized.
                Using cmd As New SqlCommand(
                    "IF NOT EXISTS (SELECT 1 FROM sys.server_principals WHERE name = @loginName) " &
                    "BEGIN " &
                    "DECLARE @stmt nvarchar(max) = N'CREATE LOGIN ' + QUOTENAME(@loginName) + N' WITH PASSWORD = N''' + REPLACE(@password, '''', '''''') + N''', CHECK_POLICY = ON, CHECK_EXPIRATION = OFF'; " &
                    "EXEC sp_executesql @stmt; " &
                    "END " &
                    "ELSE " &
                    "BEGIN " &
                    "DECLARE @reset nvarchar(max) = N'ALTER LOGIN ' + QUOTENAME(@loginName) + N' WITH PASSWORD = N''' + REPLACE(@password, '''', '''''') + N''' UNLOCK'; " &
                    "EXEC sp_executesql @reset; " &
                    "END", conn)
                    cmd.Parameters.AddWithValue("@loginName", loginName)
                    cmd.Parameters.AddWithValue("@password", password)
                    cmd.ExecuteNonQuery()
                End Using
            End Using

            Dim appDb As New SqlConnectionStringBuilder(adminConnectionString) With {.InitialCatalog = AppDatabaseName}
            Using conn As New SqlConnection(appDb.ConnectionString)
                conn.Open()
                Using cmd As New SqlCommand(
                    "IF NOT EXISTS (SELECT 1 FROM sys.database_principals WHERE name = @loginName) " &
                    "BEGIN " &
                    "DECLARE @stmt nvarchar(max) = N'CREATE USER ' + QUOTENAME(@loginName) + N' FOR LOGIN ' + QUOTENAME(@loginName); " &
                    "EXEC sp_executesql @stmt; " &
                    "END " &
                    "IF NOT EXISTS (SELECT 1 FROM sys.database_role_members drm " &
                    "JOIN sys.database_principals r ON r.principal_id = drm.role_principal_id " &
                    "JOIN sys.database_principals m ON m.principal_id = drm.member_principal_id " &
                    "WHERE r.name = N'db_owner' AND m.name = @loginName) " &
                    "BEGIN " &
                    "DECLARE @role nvarchar(max) = N'ALTER ROLE [db_owner] ADD MEMBER ' + QUOTENAME(@loginName); " &
                    "EXEC sp_executesql @role; " &
                    "END", conn)
                    cmd.Parameters.AddWithValue("@loginName", loginName)
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        ''' <summary>
        ''' Seats reach the server over TCP on 1433, but a fresh Express install listens
        ''' on shared memory only. Writes the TCP settings through xp_instance_regwrite,
        ''' the same instance-relative mechanism the Mixed Mode flip uses, so they land
        ''' under the instance's real registry key whatever its name is. Takes effect
        ''' when NetworkPrep restarts the SQL service.
        ''' </summary>
        Public Shared Sub EnsureTcpEnabled(adminConnectionString As String)
            Dim master As New SqlConnectionStringBuilder(adminConnectionString) With {.InitialCatalog = "master"}
            Using conn As New SqlConnection(master.ConnectionString)
                conn.Open()
                Using cmd As New SqlCommand(
                    "EXEC xp_instance_regwrite N'HKEY_LOCAL_MACHINE', N'Software\Microsoft\MSSQLServer\MSSQLServer\SuperSocket.NetLib\Tcp', N'Enabled', REG_DWORD, 1; " &
                    "EXEC xp_instance_regwrite N'HKEY_LOCAL_MACHINE', N'Software\Microsoft\MSSQLServer\MSSQLServer\SuperSocket.NetLib\Tcp', N'TcpPort', REG_SZ, N'1433'; " &
                    "EXEC xp_instance_regwrite N'HKEY_LOCAL_MACHINE', N'Software\Microsoft\MSSQLServer\MSSQLServer\SuperSocket.NetLib\Tcp', N'TcpDynamicPorts', REG_SZ, N''", conn)
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Private Shared Sub ExecuteScript(scriptName As String, catalog As String, adminConnectionString As String)
            ExecuteBatches(scriptName, ReadEmbeddedScript(scriptName), catalog, adminConnectionString)
        End Sub

        Private Shared Sub ExecuteBatches(scriptName As String, sql As String, catalog As String, adminConnectionString As String)
            Dim builder As New SqlConnectionStringBuilder(adminConnectionString) With {.InitialCatalog = catalog}
            Using conn As New SqlConnection(builder.ConnectionString)
                conn.Open()
                For Each batch As String In SplitBatches(sql)
                    If String.IsNullOrWhiteSpace(batch) Then Continue For
                    Try
                        Using cmd As New SqlCommand(batch, conn) With {.CommandTimeout = 60}
                            cmd.ExecuteNonQuery()
                        End Using
                    Catch ex As Exception
                        Throw New Exception(String.Format("Schema script {0} failed in catalog {1}: {2}", scriptName, catalog, ex.Message), ex)
                    End Try
                Next
            End Using
        End Sub

        Private Shared Function ReadEmbeddedScript(scriptName As String) As String
            Dim stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(ResourcePrefix & scriptName)
            If stream Is Nothing Then Throw New InvalidOperationException("Embedded schema script missing from the exe: " & ResourcePrefix & scriptName)
            Using stream
                Using reader As New System.IO.StreamReader(stream)
                    Return reader.ReadToEnd()
                End Using
            End Using
        End Function
    End Class
End Namespace
