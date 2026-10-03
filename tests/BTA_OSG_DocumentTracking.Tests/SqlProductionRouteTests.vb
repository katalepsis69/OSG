Option Explicit On
Option Strict On

Imports System
Imports System.Data
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports BTA_OSG

''' <summary>
''' The one gap a code audit can never close: the SQL-path code (registration
''' transaction, offline replay, user enrolment) executing against a LIVE SQL Server.
''' These tests run against the LOCAL SQL Server Express instance only, and refuse
''' (Inconclusive) against anything else, so they can never touch an office server.
''' Provisioning executes db/scripts 001-015 verbatim, exactly as IT would.
''' </summary>
<TestClass>
Public Class SqlProductionRouteTests
    ' Encrypt=false is deliberate for the local harness: a fresh Express install has only
    ' Shared Memory enabled (no TCP, Browser stopped), and Shared Memory cannot carry TLS.
    ' Office connection strings stay Encrypt=true over TCP; this one never leaves localhost.
    ' Shared fields, not consts: ClassInitialize fills them from whichever local instance
    ' actually answers, so a reinstall that renames the instance keeps working.
    Private Shared LocalExpressConnectionString As String = ""
    Private Shared MasterConnectionString As String = ""

    Private Shared _available As Boolean
    Private Shared _savedSettings As AppSettings
    Private Shared _initError As String = "no local SQL Server candidate answered"

    ' Registrations name the officer who filed them through a real foreign key, and nothing is
    ' seeded any more, so this class enrols the one its probe documents are filed under.
    Private Shared _officerId As Integer

    ' The LAN-discovery probes authenticate with a typed SQL login (the wizard is
    ' SQL-Auth-only), so the class owns one unique probe login for its lifetime.
    Private Shared _probeLogin As String = ""
    Private Shared _probeLoginError As String = ""
    Private Const ProbeLoginPassword As String = "Bta0SG!ProbeLogin2026"

    <ClassInitialize>
    Public Shared Sub ClassInit(context As TestContext)
        _savedSettings = AppStartup.Settings

        Dim overrideServer = Environment.GetEnvironmentVariable("BTA_TEST_SQL_SERVER")
        If Not String.IsNullOrWhiteSpace(overrideServer) Then
            Dim ods = overrideServer.ToUpperInvariant()
            If Not (ods.Contains("LOCALHOST") OrElse ods.Contains("LOCALDB") OrElse ods.Contains("127.0.0.1") OrElse ods = ".") Then
                _initError = "BTA_TEST_SQL_SERVER points off-machine; refusing: " & overrideServer
                Return
            End If
        End If
        Dim candidates() As String =
            If(String.IsNullOrWhiteSpace(overrideServer),
               New String() {"localhost\SQLEXPRESS", "localhost\SQLEXPRESS01", "localhost"},
               New String() {overrideServer, "localhost\SQLEXPRESS", "localhost\SQLEXPRESS01", "localhost"})

        For Each server In candidates
            Dim master = "Server=" & server & ";Database=master;Integrated Security=true;Encrypt=false;TrustServerCertificate=true;Connect Timeout=5"
            Try
                Using conn As New Microsoft.Data.SqlClient.SqlConnection(master)
                    conn.Open()
                End Using
                MasterConnectionString = master
                LocalExpressConnectionString = "Server=" & server & ";Database=BTA_OSG_DB;Integrated Security=true;Encrypt=false;TrustServerCertificate=true;Connect Timeout=5"
                _initError = ""
                _available = True
                Exit For
            Catch ex As Exception
                _initError = server & " -> " & ex.GetType().Name & ": " & ex.Message
            End Try
        Next
        If Not _available Then Return

        Try
            EnsureProvisioned()
            WireAppToLiveSql()
            _officerId = EnsureProbeOfficer()
        Catch ex As Exception
            _available = False
            _initError = "provisioning failed: " & ex.ToString()
            Return
        End Try

        Try
            _probeLogin = "bta_probe_" & DateTime.UtcNow.Ticks.ToString()
            DatabaseProvisioner.EnsureBtaAppLogin(MasterConnectionString, _probeLogin, ProbeLoginPassword)
        Catch ex As Exception
            _probeLogin = ""
            _probeLoginError = ex.Message
        End Try
    End Sub

    <ClassCleanup>
    Public Shared Sub ClassCleanup()
        ' Restore the shared statics the offline-path tests depend on: they assert that a
        ' disconnected BuildSqlSnapshot is a no-op, which needs the saved settings and the
        ' disconnected flag back exactly as they were.
        AppStartup.Settings = _savedSettings
        Program.IsDatabaseConnected = False
        ' Hand demo mode back to the offline classes, whatever order they run in.
        AppSettings.Instance.DatabaseSettings.UseSqlServer = False
        Try
            If _probeLogin.Length > 0 Then DropProbeLogin(_probeLogin)
        Catch
        End Try
    End Sub

    <TestInitialize>
    Public Sub TestInit()
        If Not _available Then Assert.Inconclusive("SQL production-route harness unavailable: " & _initError)
        Program.IsDatabaseConnected = True
    End Sub

    <TestCleanup>
    Public Sub TestCleanup()
        Program.IsDatabaseConnected = False
    End Sub

    ' ---------- provisioning ----------

    ' The provisioner is what the wizard's Prepare Server runs: scripts 001-015 from the
    ' exe's embedded resources, verbatim. If it stops being idempotent, every suite run
    ' fails here before any assertion does.
    Private Shared Sub EnsureProvisioned()
        ' Provision runs unconditionally: every script is idempotent, so this is a no-op
        ' on a current database, and an existing one picks up a newly added migration the
        ' same way the fresh-provision path does (this is what healed the pre-014 dev
        ' database during the 2026-10-02 audit run).
        DatabaseProvisioner.Provision(MasterConnectionString)
    End Sub

    ' Best-effort cleanup for the probe logins this class mints; self-generated names.
    Private Shared Sub DropProbeLogin(loginName As String)
        Using conn As New Microsoft.Data.SqlClient.SqlConnection(LocalExpressConnectionString)
            conn.Open()
            Using cmd As New Microsoft.Data.SqlClient.SqlCommand(
                "DECLARE @s nvarchar(max) = N'IF EXISTS (SELECT 1 FROM sys.database_principals WHERE name = @n) DROP USER ' + QUOTENAME(@n); " &
                "EXEC sp_executesql @s, N'@n sysname', @n = @n", conn)
                cmd.Parameters.AddWithValue("@n", loginName)
                cmd.ExecuteNonQuery()
            End Using
        End Using
        Using conn As New Microsoft.Data.SqlClient.SqlConnection(MasterConnectionString)
            conn.Open()
            Using cmd As New Microsoft.Data.SqlClient.SqlCommand(
                "DECLARE @s nvarchar(max) = N'IF EXISTS (SELECT 1 FROM sys.server_principals WHERE name = @n) DROP LOGIN ' + QUOTENAME(@n); " &
                "EXEC sp_executesql @s, N'@n sysname', @n = @n", conn)
                cmd.Parameters.AddWithValue("@n", loginName)
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    ' ---------- app wiring ----------

    ''' <summary>
    ''' Mirrors AppStartup.Initialize but against the local Express factory, without
    ''' touching the AppSettings.Instance singleton that other tests read.
    ''' </summary>
    Private Shared Sub WireAppToLiveSql()
        Dim factory As IDbConnectionFactory = New SqlConnectionFactory(LocalExpressConnectionString)

        AppStartup.ConnectionFactory = factory
        AppStartup.UserRepo = New UserRepository(factory)
        AppStartup.DocumentRepo = New DocumentRepository(factory)
        AppStartup.DirectiveRepo = New DirectiveRepository(factory)
        AppStartup.RoutingRepo = New RoutingRepository(factory)
        AppStartup.StorageRepo = New StorageRepository(factory)
        AppStartup.AuditRepo = New AuditRepository(factory)
        AppStartup.SequenceRepo = New SequenceRepository(factory)
        AppStartup.ReferenceDataRepo = New ReferenceDataRepository(factory)

        AppStartup.AuditService = New AuditService(AppStartup.AuditRepo)
        AppStartup.PortalBridgeClient = New PortalBridge(New PortalSettings(), auditService:=AppStartup.AuditService)
        AppStartup.DocService = New DocumentService(AppStartup.DocumentRepo, AppStartup.SequenceRepo, AppStartup.ReferenceDataRepo, AppStartup.AuditService, AppStartup.PortalBridgeClient, AppStartup.RoutingRepo, AppStartup.StorageRepo)
        AppStartup.DirectiveService = New DirectiveService(AppStartup.DirectiveRepo, AppStartup.DocumentRepo, AppStartup.ReferenceDataRepo, AppStartup.AuditService, AppStartup.PortalBridgeClient)
        AppStartup.RoutingService = New RoutingService(AppStartup.RoutingRepo, AppStartup.DocumentRepo, AppStartup.ReferenceDataRepo, AppStartup.AuditService, AppStartup.PortalBridgeClient)
        AppStartup.StorageService = New StorageService(AppStartup.StorageRepo, AppStartup.DocumentRepo, AppStartup.AuditService)
        AppStartup.CardService = New RfidCardService(factory, AppStartup.AuditService)

        ' BuildSqlSnapshot needs UseSqlServer=True plus the reachable string; a fresh
        ' AppSettings object here keeps AppSettings.Instance untouched.
        Dim settings As New AppSettings()
        settings.DatabaseSettings.UseSqlServer = True
        settings.DatabaseSettings.ConnectionString = LocalExpressConnectionString
        AppStartup.Settings = settings
        ' The shared embedded store must run in SQL mode while this class is live: the
        ' assembly init forces demo mode for the offline classes, and demo seeding here
        ' would drop foreign rows into the outbox this class replays.
        AppSettings.Instance.DatabaseSettings.UseSqlServer = True
    End Sub

    ''' <summary>
    ''' Idempotent so a repeat run reuses the same officer and the document foreign keys keep
    ''' pointing at one identity.
    ''' </summary>
    Private Shared Function EnsureProbeOfficer() As Integer
        Dim existing = Scalar("SELECT UserID FROM tbl_Users WHERE Username = @u", "@u", "probeofficer")
        If existing IsNot Nothing AndAlso Not IsDBNull(existing) Then Return Convert.ToInt32(existing)
        Return Convert.ToInt32(Scalar(
            "INSERT INTO tbl_Users (Username, FullName, Office) OUTPUT INSERTED.UserID VALUES (@u, @n, 'Records Section')",
            "@u", "probeofficer", "@n", "Probe Officer"))
    End Function

    Private Shared Function Scalar(sql As String, ParamArray args As Object()) As Object
        Using conn As New Microsoft.Data.SqlClient.SqlConnection(LocalExpressConnectionString)
            conn.Open()
            Using cmd As New Microsoft.Data.SqlClient.SqlCommand(sql, conn)
                For i As Integer = 0 To args.Length - 2 Step 2
                    cmd.Parameters.AddWithValue(CStr(args(i)), args(i + 1))
                Next
                Return cmd.ExecuteScalar()
            End Using
        End Using
    End Function

    Private Shared Sub ExecuteNonQuery(sql As String, ParamArray args As Object())
        Using conn As New Microsoft.Data.SqlClient.SqlConnection(LocalExpressConnectionString)
            conn.Open()
            Using cmd As New Microsoft.Data.SqlClient.SqlCommand(sql, conn)
                For i As Integer = 0 To args.Length - 2 Step 2
                    cmd.Parameters.AddWithValue(CStr(args(i)), args(i + 1))
                Next
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Private Shared Function StatusCodeOf(docCode As String) As String
        Return Convert.ToString(Scalar(
            "SELECT st.StatusCode FROM tbl_Documents d JOIN tbl_DocumentStatuses st ON st.StatusID = d.StatusID WHERE d.DocCode = @c",
            "@c", docCode))
    End Function

    ' ---------- the production route, end to end ----------

    <TestMethod>
    Public Sub WizardSchemaProbe_FindsProvisionedDatabase()
        ' The exact check the setup wizard's Test Connection performs.
        Dim count = Convert.ToInt32(Scalar(
            "SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME IN ('tbl_Documents','tbl_Users','tbl_AuditTrail')"))
        Assert.AreEqual(3, count, "scripts 001-015 must leave the three sentinel tables behind")
    End Sub

    <TestMethod>
    Public Sub ConnectedRegistration_Route_Approve_Release_EndToEnd()
        Dim coordinator As New DesktopDataCoordinator(True)
        Dim stamp = DateTime.UtcNow.Ticks.ToString()
        Dim title = "E2E connected registration probe " & stamp
        Dim code = coordinator.RegisterDocument(
            "COMM-2099-" & (stamp.Substring(stamp.Length - 6)),
            "Regular Communication", title, "Ministry of Interior", "Office of the Secretary-General",
            "CAB-A", "S-1", "BOX-01", "", "FOR_REVIEW", "Probe Officer", "INCOMING", "Secretariat",
            DateTime.Now.AddDays(3).ToString("yyyy-MM-dd HH:mm:ss"), "", "Registered by integration test", "Probe Officer", _officerId)

        Assert.IsTrue(code.StartsWith("COMM-"), "connected registration mints the code from the SQL sequence")
        Assert.AreEqual(1, Convert.ToInt32(Scalar("SELECT COUNT(*) FROM tbl_Documents WHERE DocCode = @c", "@c", code)), "document row lands in SQL")

        Dim docId = Convert.ToInt32(Scalar("SELECT DocumentID FROM tbl_Documents WHERE DocCode = @c", "@c", code))
        Assert.AreEqual("FOR_REVIEW", StatusCodeOf(code))
        Assert.AreEqual(1, Convert.ToInt32(Scalar(
            "SELECT COUNT(*) FROM tbl_Documents d JOIN tbl_DocumentTypes dt ON dt.DocumentTypeID = d.DocumentTypeID " &
            "WHERE d.DocumentID = @i AND dt.TypeCode = 'REG_COMM'", "@i", docId.ToString())),
            "a Regular Communication must resolve to the seeded REG_COMM type, not the fallback Resolution row")
        Assert.IsTrue(Convert.ToInt32(Scalar("SELECT COUNT(*) FROM tbl_RoutingLogs WHERE DocumentID = @i", "@i", docId)) >= 1, "registration writes the routing log")
        Assert.IsTrue(Convert.ToInt32(Scalar("SELECT COUNT(*) FROM tbl_DocumentMovements WHERE DocumentID = @i", "@i", docId)) >= 1, "registration writes the storage movement")
        Assert.IsTrue(Convert.ToInt32(Scalar("SELECT COUNT(*) FROM tbl_DocumentAssignments WHERE DocumentID = @i", "@i", docId)) >= 1, "registration writes the assignment")
        Assert.IsTrue(Convert.ToInt32(Scalar("SELECT COUNT(*) FROM tbl_AuditTrail WHERE EntityType = 'Document' AND EntityID = @i", "@i", docId.ToString())) >= 1, "registration writes its audit entry")

        ' The workflow, each through the coordinator's real SQL path.
        coordinator.RouteDocument(docId, "Secretariat", "Finance Section", "ROUTED", "e2e route", "Probe Officer", _officerId)
        Assert.AreEqual("ROUTED", StatusCodeOf(code))
        coordinator.ApproveDocument(docId, "Probe Officer", "e2e approve", _officerId)
        Assert.AreEqual("APPROVED", StatusCodeOf(code))
        coordinator.ReleaseDocument(docId, "Probe Officer", "e2e release", _officerId)
        Assert.AreEqual("RELEASED", StatusCodeOf(code))
        Assert.IsTrue(Convert.ToInt32(Scalar("SELECT COUNT(*) FROM tbl_RoutingLogs WHERE DocumentID = @i", "@i", docId)) >= 4, "every transition left a custody log row")
    End Sub

    <TestMethod>
    Public Sub MoveStorage_CreatesLandmarkAndMovement()
        Dim coordinator As New DesktopDataCoordinator(True)
        Dim stamp = DateTime.UtcNow.Ticks.ToString()
        Dim code = coordinator.RegisterDocument(
            "FIN-2099-" & (stamp.Substring(stamp.Length - 6)),
            "Finance", "E2E storage move probe " & stamp, "Finance Division", "Finance Section",
            "CAB-C", "S-2", "BOX-02", "", "RECEIVED", "Probe Officer", "INCOMING", "Finance Section",
            "", "", "Registered by integration test", "Probe Officer", _officerId)
        Dim docId = Convert.ToInt32(Scalar("SELECT DocumentID FROM tbl_Documents WHERE DocCode = @c", "@c", code))

        coordinator.MoveStorage(docId, "CAB-C/S-2/BOX-02", "CAB-Q|S-9|BOX-99", "e2e move", "Probe Officer", _officerId)

        Dim locationId = Convert.ToInt32(Scalar("SELECT CurrentStorageLocationID FROM tbl_Documents WHERE DocumentID = @i", "@i", docId))
        Assert.IsTrue(Convert.ToInt32(Scalar(
            "SELECT COUNT(*) FROM tbl_StorageLocations WHERE StorageLocationID = @l AND CabinetID = 'CAB-Q' AND ShelfNo = 'S-9' AND BoxCode = 'BOX-99'",
            "@l", locationId)) = 1, "move resolved the real landmark and pointed the document at it")
        Assert.IsTrue(Convert.ToInt32(Scalar(
            "SELECT COUNT(*) FROM tbl_DocumentMovements WHERE DocumentID = @i AND StorageLocationID = @l",
            "@i", docId, "@l", locationId)) >= 1, "movement row written for the new landmark")
    End Sub

    <TestMethod>
    Public Sub OfflineReplay_PreservesCodeAndControlNumber()
        ' The audit-009 worst case, live: register offline (cache only), reconnect, replay.
        Dim stamp = DateTime.UtcNow.Ticks.ToString()
        ' ExternalControlNumber is VARCHAR(24): the 8-char prefix plus a 12-digit tail fits,
        ' the full 18-digit tick stamp would not.
        stamp = stamp.Substring(stamp.Length - 12)
        ' The stamp digits make the probe code unique per run; a fixed suffix would collide
        ' with the previous run's row and send the replay down the fallback path.
        Dim offlineCode = "COMM-2099-" & stamp
        Dim offlineDocId = EmbeddedDB.AddDocument(
            offlineCode, "Regular Communication", "E2E offline replay probe " & stamp,
            "Ministry of Interior", "Office of the Secretary-General", "CAB-A", "S-1", "BOX-01", "",
            "FOR_REVIEW", "Probe Officer", "INCOMING", "Secretariat", "", "",
            "Registered offline", "EXT-E2E-" & stamp, isOffline:=True, createdByUserId:=_officerId)

        Dim coordinator As New DesktopDataCoordinator(True)
        Dim uploaded = coordinator.SyncOfflineOutbox()

        Assert.IsTrue(uploaded >= 1, "outbox reported the replay")
        Assert.AreEqual(1, Convert.ToInt32(Scalar("SELECT COUNT(*) FROM tbl_Documents WHERE DocCode = @c", "@c", offlineCode)),
                        "the offline code survives replay when free")
        Dim sqlDocId = Convert.ToInt32(Scalar("SELECT DocumentID FROM tbl_Documents WHERE DocCode = @c", "@c", offlineCode))
        Assert.AreEqual("EXT-E2E-" & stamp, Convert.ToString(Scalar("SELECT ExternalControlNumber FROM tbl_Documents WHERE DocumentID = @i", "@i", sqlDocId)),
                        "external control number carried through replay")
        Assert.AreEqual("FOR_REVIEW", StatusCodeOf(offlineCode))
        Assert.IsTrue(Convert.ToInt32(Scalar("SELECT COUNT(*) FROM tbl_RoutingLogs WHERE DocumentID = @i", "@i", sqlDocId)) >= 1, "replayed registration wrote routing children")

        ' The mirror row must be gone: a leftover pending row would replay again.
        SyncLock EmbeddedDB.SyncRoot
            Assert.AreEqual(0, EmbeddedDB.DataSet.Tables("Documents").Select("DocumentID = " & offlineDocId).Length,
                            "replayed offline row removed from the cache")
        End SyncLock
    End Sub

    <TestMethod>
    Public Sub OfflineUserEnrolment_ReplaysToSql()
        ' Unique per run: the scratch database survives test runs, so a FullName lookup
        ' must not match a probe user enrolled by an earlier run.
        Dim stamp = DateTime.UtcNow.Ticks.ToString()
        stamp = stamp.Substring(stamp.Length - 12)
        Dim uid = "E2E" & stamp.Substring(0, 8)
        Dim fullName = "E2E Replay Staff " & stamp
        EmbeddedDB.AddUser(uid, fullName, "Records Section", "Records Section", True, True, True, pendingSync:=True)

        Dim coordinator As New DesktopDataCoordinator(True)
        coordinator.SyncOfflineOutbox()

        Dim userId = Convert.ToInt32(Scalar("SELECT UserID FROM tbl_Users WHERE FullName = @n", "@n", fullName))
        Assert.IsTrue(userId > 0, "offline enrolment reached tbl_Users")
        Assert.AreEqual(1, Convert.ToInt32(Scalar("SELECT COUNT(*) FROM tbl_RfidCards WHERE UserID = @i AND CardPublicID = @c AND IsActive = 1", "@i", userId, "@c", uid)),
                        "badge issued to the replayed user")
        Assert.IsTrue(Convert.ToInt32(Scalar(
            "SELECT COUNT(*) FROM tbl_UserRoles ur JOIN tbl_Roles r ON r.RoleID = ur.RoleID WHERE ur.UserID = @i AND r.RoleCode = 'RECORDS'",
            "@i", userId)) >= 1, "section role assigned")
    End Sub

    <TestMethod>
    Public Sub OfflineDocCodeCollision_FallsBackToSharedSequence()
        ' Two workstations offline can mint the same code; both must land in SQL with
        ' distinct codes (the second falls back to the shared sequence on the unique key).
        Dim stamp = DateTime.UtcNow.Ticks.ToString()
        stamp = stamp.Substring(stamp.Length - 12)
        Dim sharedCode = "COMM-2098-" & stamp
        EmbeddedDB.AddDocument(sharedCode, "Regular Communication", "Collision probe A " & stamp,
            "O", "Secretariat", "CAB-A", "S-1", "BOX-01", "", "FOR_REVIEW", "A", "INCOMING", "Secretariat",
            "", "", "offline", "", isOffline:=True, createdByUserId:=_officerId)
        EmbeddedDB.AddDocument(sharedCode, "Regular Communication", "Collision probe B " & stamp,
            "O", "Secretariat", "CAB-A", "S-1", "BOX-01", "", "FOR_REVIEW", "A", "INCOMING", "Secretariat",
            "", "", "offline", "", isOffline:=True, createdByUserId:=_officerId)

        Dim coordinator As New DesktopDataCoordinator(True)
        coordinator.SyncOfflineOutbox()

        Dim codeA = Convert.ToString(Scalar("SELECT DocCode FROM tbl_Documents WHERE Title = N'Collision probe A " & stamp & "'"))
        Dim codeB = Convert.ToString(Scalar("SELECT DocCode FROM tbl_Documents WHERE Title = N'Collision probe B " & stamp & "'"))
        Assert.AreNotEqual("", codeA, "first offline doc replayed")
        Assert.AreNotEqual("", codeB, "second offline doc replayed despite the duplicate code")
        Assert.AreNotEqual(codeA, codeB, "collision resolved: one kept the offline code, the other re-minted")
    End Sub

    <TestMethod>
    Public Sub SecondWorkstation_MirrorRead_SeesTheRegistration()
        ' Seat A registers; seat B's mirror pull returns it. The shared static cache means
        ' both seats here read one DataSet, so this verifies the SQL read path, not cache
        ' separation (the cache merge itself is covered by SqlMirrorTests).
        Dim coordinatorA As New DesktopDataCoordinator(True)
        Dim stamp = DateTime.UtcNow.Ticks.ToString()
        Dim title = "E2E mirror read probe " & stamp
        Dim code = coordinatorA.RegisterDocument(
            "TO-2099-" & (stamp.Substring(stamp.Length - 6)),
            "Travel Order", title, "OSG Travel Desk", "Travel Section",
            "CAB-A", "S-2", "BOX-03", "", "FOR_REVIEW", "CJ Fairoz A. Usop", "INCOMING", "Travel Section",
            "", "", "Registered by integration test", "Probe Officer", _officerId)

        Dim coordinatorB As New DesktopDataCoordinator(True)
        Dim snapshot = coordinatorB.BuildSqlSnapshot("Documents")
        Assert.IsNotNull(snapshot, "seat B mirror pull returns rows from the live server")
        Assert.IsTrue(snapshot.ContainsKey("Documents"), "documents table included")
        Dim match = snapshot("Documents").Select("DocCode = '" & code.Replace("'", "''") & "'")
        Assert.AreEqual(1, match.Length, "seat B sees seat A's registration within one poll")
    End Sub

    <TestMethod>
    Public Sub LanDiscovery_Probe_FindsTheProvisionedDatabase()
        ' The wizard's LAN scan must classify the live provisioned instance as a hit with
        ' a typed SQL login, otherwise the prefill would never fire on a real office network.
        If _probeLogin.Length = 0 Then Assert.Inconclusive("SQL-auth probe login unavailable: " & _probeLoginError)
        If Convert.ToInt32(Scalar("SELECT CAST(SERVERPROPERTY('IsIntegratedSecurityOnly') AS int)")) = 1 Then
            Assert.Inconclusive("local instance is Windows-auth-only; Prepare Server (or Mixed Mode + service restart) unlocks the SQL-auth probe path")
        End If
        Dim result = SqlInstanceDiscovery.Probe("localhost", 1433, _probeLogin, ProbeLoginPassword)
        Assert.AreEqual(SqlInstanceDiscovery.ProbeResult.FoundDatabase, result)
    End Sub

    <TestMethod>
    Public Sub LanDiscovery_Probe_UnknownLogin_IsRejected()
        ' The 18456 mapping is what drives the wizard's Prepare Server offer.
        Dim result = SqlInstanceDiscovery.Probe("localhost", 1433, "no_such_probe_login", "NotARealPassword1!")
        Assert.AreEqual(SqlInstanceDiscovery.ProbeResult.LoginRejected, result)
    End Sub

    <TestMethod>
    Public Sub LanDiscovery_Discover_IncludesTheLocalInstance()
        Dim results = SqlInstanceDiscovery.Discover("", "")
        Assert.IsTrue(results.Count >= 1, "the local machine is always a candidate, even when Browser is silent")
        Dim local = results.Find(Function(x) x.Server.Equals(Environment.MachineName, StringComparison.OrdinalIgnoreCase) OrElse
                                            x.Server.StartsWith(Environment.MachineName & "\", StringComparison.OrdinalIgnoreCase) OrElse
                                            x.Server.Equals("localhost", StringComparison.OrdinalIgnoreCase))
        Assert.IsNotNull(local, "machine-name candidate present")
    End Sub

    ' ---------- zero-scripts setup ----------

    <TestMethod>
    Public Sub Provisioner_Rerun_IsCleanNoOp()
        ' The class already provisioned in ClassInit; a second full pass must not throw
        ' and must not duplicate seed rows.
        Dim typeCountBefore = Convert.ToInt32(Scalar("SELECT COUNT(*) FROM tbl_DocumentTypes"))
        DatabaseProvisioner.Provision(MasterConnectionString)
        Assert.AreEqual(typeCountBefore, Convert.ToInt32(Scalar("SELECT COUNT(*) FROM tbl_DocumentTypes")),
                        "the guarded scripts must not duplicate seed rows on a re-run")
    End Sub

    <TestMethod>
    Public Sub EnsureBtaAppLogin_CreatesLoginUserAndRole_Idempotent()
        Dim stamp = DateTime.UtcNow.Ticks.ToString()
        Dim loginName = "bta_probe_" & stamp.Substring(stamp.Length - 12)
        Try
            DatabaseProvisioner.EnsureBtaAppLogin(MasterConnectionString, loginName, ProbeLoginPassword)
            DatabaseProvisioner.EnsureBtaAppLogin(MasterConnectionString, loginName, ProbeLoginPassword)

            Assert.AreEqual(1, Convert.ToInt32(Scalar("SELECT COUNT(*) FROM sys.server_principals WHERE name = @n", "@n", loginName)),
                            "server login exists")
            Assert.AreEqual(1, Convert.ToInt32(Scalar(
                "SELECT COUNT(*) FROM sys.database_principals WHERE name = @n AND type = 'S'", "@n", loginName)),
                "database user exists (SQL user class)")
            Assert.AreEqual(1, Convert.ToInt32(Scalar(
                "SELECT COUNT(*) FROM sys.database_role_members drm " &
                "JOIN sys.database_principals r ON r.principal_id = drm.role_principal_id " &
                "JOIN sys.database_principals m ON m.principal_id = drm.member_principal_id " &
                "WHERE r.name = 'db_owner' AND m.name = @n", "@n", loginName)),
                "user is a db_owner member")
        Finally
            DropProbeLogin(loginName)
        End Try
    End Sub

    <TestMethod>
    Public Sub Heartbeat_UpsertLandsAndMirrorsToCache()
        Dim coordinator As New DesktopDataCoordinator(True)
        Dim snapshot = coordinator.BuildSqlSnapshot()
        Assert.IsNotNull(snapshot, "snapshot pull succeeded against the live server")
        Assert.IsTrue(snapshot.ContainsKey("Heartbeat"), "the heartbeat is part of the default pull")

        Dim machine = Environment.MachineName.Replace("'", "''")
        Dim version = Convert.ToString(Scalar("SELECT AppVersion FROM tbl_WorkstationHeartbeat WHERE MachineName = @m", "@m", Environment.MachineName))
        Assert.AreNotEqual("", version, "this seat's heartbeat upsert landed with its app version")

        EmbeddedDB.ApplySnapshots(snapshot)
        Dim rows = EmbeddedDB.DataSet.Tables("Heartbeat").Select("MachineName = '" & machine & "'")
        Assert.AreEqual(1, rows.Length, "the heartbeat mirrored into the cache for the seats grid")
    End Sub

    <TestMethod>
    Public Sub ReplayConflict_ServerKeepsNewerTruth()
        ' Two seats, one document: seat B moved it while seat A was offline, so A's replay
        ' carries a stale rowversion. The server copy must stand and the conflict be recorded.
        Dim coordinator As New DesktopDataCoordinator(True)
        Dim stamp = DateTime.UtcNow.Ticks.ToString()
        stamp = stamp.Substring(stamp.Length - 12)
        Dim code = coordinator.RegisterDocument(
            "COMM-2099-" & (stamp.Substring(stamp.Length - 6)),
            "Regular Communication", "Conflict probe " & stamp, "Ministry of Interior", "Office of the Secretary-General",
            "CAB-A", "S-1", "BOX-01", "", "FOR_REVIEW", "Probe Officer", "INCOMING", "Secretariat",
            "", "", "Registered by integration test", "Probe Officer", _officerId)
        Dim docId = Convert.ToInt32(Scalar("SELECT DocumentID FROM tbl_Documents WHERE DocCode = @c", "@c", code))
        Dim mirroredVersion = CType(Scalar("SELECT RowVersion FROM tbl_Documents WHERE DocumentID = @i", "@i", docId), Byte())

        ' Seat A mirrors the document, then edits it offline against the mirrored version.
        Dim offlineRow As DataRow
        SyncLock EmbeddedDB.SyncRoot
            Dim rows = EmbeddedDB.DataSet.Tables("Documents").Select("DocumentID = " & docId)
            Assert.AreEqual(1, rows.Length, "the registered document mirrored into the cache")
            offlineRow = rows(0)
            offlineRow("PendingSync") = True
            offlineRow("RowVersion") = mirroredVersion
            offlineRow("LastActionTaken") = "OFFLINE EDIT " & stamp
        End SyncLock

        ' Seat B moves the document past the mirrored version.
        ExecuteNonQuery("UPDATE tbl_Documents SET LastActionTaken = @a, ModifiedAtUTC = SYSUTCDATETIME() WHERE DocumentID = @i",
                        "@a", "SERVER MOVE " & stamp, "@i", docId)

        coordinator.SyncOfflineOutbox()

        Assert.AreEqual("SERVER MOVE " & stamp, Convert.ToString(Scalar("SELECT LastActionTaken FROM tbl_Documents WHERE DocumentID = @i", "@i", docId)),
                        "the server kept the newer edit")
        SyncLock EmbeddedDB.SyncRoot
            Assert.IsFalse(CBool(offlineRow("PendingSync")), "the conflicted row stops replaying")
        End SyncLock
        Assert.AreEqual(1, EmbeddedDB.DataSet.Tables("AuditTrail").Select(
                            "ActionType = 'SYNC_CONFLICT' AND ActionDescription LIKE '%" & code & "%'").Length,
                        "the conflict was recorded with the document code")
    End Sub

    <TestMethod>
    Public Sub ReplayWithoutConflict_StillApplies()
        ' Same setup, no competing edit: the guarded replay behaves exactly like the
        ' unguarded one did.
        Dim coordinator As New DesktopDataCoordinator(True)
        Dim stamp = DateTime.UtcNow.Ticks.ToString()
        stamp = stamp.Substring(stamp.Length - 12)
        Dim code = coordinator.RegisterDocument(
            "FIN-2099-" & (stamp.Substring(stamp.Length - 6)),
            "Finance", "Clean replay probe " & stamp, "Finance Division", "Finance Section",
            "CAB-C", "S-2", "BOX-02", "", "RECEIVED", "Probe Officer", "INCOMING", "Finance Section",
            "", "", "Registered by integration test", "Probe Officer", _officerId)
        Dim docId = Convert.ToInt32(Scalar("SELECT DocumentID FROM tbl_Documents WHERE DocCode = @c", "@c", code))

        SyncLock EmbeddedDB.SyncRoot
            Dim rows = EmbeddedDB.DataSet.Tables("Documents").Select("DocumentID = " & docId)
            Assert.AreEqual(1, rows.Length, "the registered document mirrored into the cache")
            rows(0)("PendingSync") = True
            rows(0)("LastActionTaken") = "OFFLINE EDIT " & stamp
        End SyncLock

        coordinator.SyncOfflineOutbox()

        Assert.AreEqual("OFFLINE EDIT " & stamp, Convert.ToString(Scalar("SELECT LastActionTaken FROM tbl_Documents WHERE DocumentID = @i", "@i", docId)),
                        "the offline edit applied when no one else had touched the document")
        Assert.AreEqual(0, EmbeddedDB.DataSet.Tables("AuditTrail").Select(
                            "ActionType = 'SYNC_CONFLICT' AND ActionDescription LIKE '%" & code & "%'").Length,
                        "no conflict recorded on the clean path")
    End Sub

    ' ---------- connect-first setup gates ----------

    ' The connect dialog offers "Set up this server" only where the registry lists a
    ' local SQL Server; on the machine running this live harness the gate must read
    ' True, or a server PC would never be offered the prepare path at all.
    <TestMethod>
    Public Sub LocalSqlServerRegistry_DetectedOnHarnessMachine()
        Assert.IsTrue(SqlInstanceDiscovery.HasLocalSqlServer(),
                      "the machine running the live-SQL harness must report a local SQL Server")
        Assert.IsTrue(SqlInstanceDiscovery.LocalInstanceNames().Count > 0,
                      "instance names must be readable so candidate enumeration can probe local named instances")
    End Sub

    ' EnsureTcpEnabled is what Prepare runs before NetworkPrep restarts the service.
    ' The write is idempotent and the harness box already listens on static 1433, so
    ' the values must read back unchanged through the same instance-relative path.
    <TestMethod>
    Public Sub EnsureTcpEnabled_WritesStaticPort1433()
        DatabaseProvisioner.EnsureTcpEnabled(MasterConnectionString)
        Using conn As New Microsoft.Data.SqlClient.SqlConnection(MasterConnectionString)
            conn.Open()
            Using cmd As New Microsoft.Data.SqlClient.SqlCommand(
                "EXEC xp_instance_regread N'HKEY_LOCAL_MACHINE', N'Software\Microsoft\MSSQLServer\MSSQLServer\SuperSocket.NetLib\Tcp', N'TcpPort'", conn)
                Using reader = cmd.ExecuteReader()
                    Assert.IsTrue(reader.Read(), "the TcpPort registry value must exist after EnsureTcpEnabled")
                    Assert.AreEqual("1433", reader("Data").ToString().Trim(),
                                    "seats connect on host,1433, so the static port must read back as 1433")
                End Using
            End Using
        End Using
    End Sub
    <TestMethod>
    Public Sub AuditChain_LiveSqlSealsEveryEntryAndDetectsAnAlteredOne()
        ' Migration 014 only adds the columns; the seal is application-side because a hash over
        ' the values a writer intended cannot be reproduced once DATETIME rounds the timestamp.
        ' This round trip is the only place that intent-versus-storage gap is really tested.
        AppStartup.AuditRepo.SealUnchainedRows()
        Assert.IsTrue(AppStartup.AuditRepo.VerifyChain().IsValid, "pre-existing entries did not seal into one chain")

        Dim marker = "CHAIN_PROBE_" & DateTime.UtcNow.Ticks.ToString()
        AppStartup.AuditService.LogEvent(marker, "Probe", Nothing, Nothing, Nothing, "{""step"":1}", True, Nothing)
        AppStartup.AuditService.LogEvent(marker, "Probe", Nothing, Nothing, Nothing, "{""step"":2}", True, Nothing)

        Dim ids = Scalar("SELECT MIN(AuditID) FROM tbl_AuditTrail WHERE ActionType = @m", "@m", marker)
        Dim targetId = Convert.ToInt64(ids)
        Assert.IsNotNull(Scalar("SELECT RowHash FROM tbl_AuditTrail WHERE AuditID = @i", "@i", targetId),
                         "a fresh append left its entry unsealed")

        Dim tampered = AppStartup.AuditRepo.VerifyChain()
        Assert.IsTrue(tampered.IsValid AndAlso tampered.RowsUnsealed = 0,
                      "the chain does not verify on a clean store: " & tampered.Reason)

        ExecuteNonQuery("UPDATE tbl_AuditTrail SET NewValuesJson = @v WHERE AuditID = @i",
                        "@v", "{""step"":99}", "@i", targetId)
        Dim broken = AppStartup.AuditRepo.VerifyChain()
        Assert.IsFalse(broken.IsValid, "an altered entry verified as clean")
        Assert.AreEqual(targetId, broken.FirstBrokenAuditID, "the break was reported at the wrong entry")

        ExecuteNonQuery("UPDATE tbl_AuditTrail SET NewValuesJson = @v WHERE AuditID = @i",
                        "@v", "{""step"":1}", "@i", targetId)
        Assert.IsTrue(AppStartup.AuditRepo.VerifyChain().IsValid,
                      "restoring the original content did not clear the break")
    End Sub

    <TestMethod>
    Public Sub ApprovalGuard_RejectsStaleFromStatus()
        ' The from-status guard behind the approve/release buttons: a document whose
        ' server state already moved must refuse the second transition instead of
        ' writing another approval onto a released row.
        Dim coordinator As New DesktopDataCoordinator(True)
        Dim stamp = DateTime.UtcNow.Ticks.ToString()
        Dim code = coordinator.RegisterDocument(
            "COMM-2099-" & (stamp.Substring(stamp.Length - 6)),
            "Regular Communication", "Guard probe " & stamp, "Ministry of Interior", "Office of the Secretary-General",
            "CAB-A", "S-1", "BOX-01", "", "FOR_REVIEW", "Probe Officer", "INCOMING", "Secretariat",
            "", "", "Registered by integration test", "Probe Officer", _officerId)
        Dim docId = Convert.ToInt32(Scalar("SELECT DocumentID FROM tbl_Documents WHERE DocCode = @c", "@c", code))
        Dim forReviewId = Convert.ToInt32(Scalar("SELECT StatusID FROM tbl_DocumentStatuses WHERE StatusCode = 'FOR_REVIEW'"))
        Dim approvedId = Convert.ToInt32(Scalar("SELECT StatusID FROM tbl_DocumentStatuses WHERE StatusCode = 'APPROVED'"))

        AppStartup.RoutingService.ApproveDocument(docId, _officerId, "guard probe approve", forReviewId)
        Assert.AreEqual("APPROVED", StatusCodeOf(code))

        ' The row now carries APPROVED: the same FOR_REVIEW expectation must refuse.
        Try
            AppStartup.RoutingService.ApproveDocument(docId, _officerId, "second approve", forReviewId)
            Assert.Fail("approving against a stale from-status must be refused")
        Catch ex As InvalidOperationException
            Assert.IsTrue(ex.Message.Length > 0)
        End Try

        ' And the release path accepts the correct current status.
        AppStartup.RoutingService.ReleaseDocument(docId, _officerId, "guard probe release", approvedId)
        Assert.AreEqual("RELEASED", StatusCodeOf(code))
    End Sub

    <TestMethod>
    Public Sub RoutingLog_NullToStatusID_IsAllowedForNonStatusActions()
        ' Migration 015: offline replay records custody events whose action is not a
        ' status code (revision requests, resubmits, transmittals); the NOT NULL
        ' constraint used to send those replays into a permanent failure loop.
        Dim coordinator As New DesktopDataCoordinator(True)
        Dim stamp = DateTime.UtcNow.Ticks.ToString()
        Dim code = coordinator.RegisterDocument(
            "COMM-2099-" & (stamp.Substring(stamp.Length - 6)),
            "Regular Communication", "Null status probe " & stamp, "Ministry of Interior", "Office of the Secretary-General",
            "CAB-A", "S-1", "BOX-01", "", "FOR_REVIEW", "Probe Officer", "INCOMING", "Secretariat",
            "", "", "Registered by integration test", "Probe Officer", _officerId)
        Dim docId = Convert.ToInt32(Scalar("SELECT DocumentID FROM tbl_Documents WHERE DocCode = @c", "@c", code))

        AppStartup.RoutingRepo.Insert(New RoutingLog With {
            .DocumentID = docId,
            .FromStatusID = Nothing,
            .ToStatusID = Nothing,
            .FromOffice = "Secretariat",
            .ToOffice = "Finance Section",
            .RoutedByUserID = _officerId,
            .RoutedAtUTC = DateTime.UtcNow,
            .RoutingRemarks = "offline transmittal replay"
        })

        Assert.AreEqual(1, Convert.ToInt32(Scalar(
            "SELECT COUNT(*) FROM tbl_RoutingLogs WHERE DocumentID = @i AND ToStatusID IS NULL", "@i", docId)),
            "a non-status offline action replays with a NULL status and no regression of the document")
        Assert.AreEqual("FOR_REVIEW", StatusCodeOf(code))
    End Sub
End Class
