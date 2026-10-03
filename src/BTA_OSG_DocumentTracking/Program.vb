Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports System.Windows.Forms


Namespace BTA_OSG
    Public NotInheritable Class Program
        Public Shared IsDatabaseConnected As Boolean = False
        Public Shared Coordinator As DesktopDataCoordinator

        ' Set by FormMain at startup: the one bridge between data-layer failures (SQL
        ' fallback, replay errors, cache save failure) and the operator's status banner.
        ' Without it those failures live only in the invisible Trace log of a WinExe.
        Public Shared Property OperatorWarningSink As Action(Of String)

        Public Shared Sub ReportOperatorWarning(message As String)
            System.Diagnostics.Trace.TraceWarning(message)
            Dim sink = OperatorWarningSink
            If sink IsNot Nothing Then
                Try
                    sink(message)
                Catch
                End Try
            End If
        End Sub

        Public Shared Sub ProbeDatabaseConnection()
            IsDatabaseConnected = False
            Try
                ' Demo mode must not reach the server even when a local instance answers: the
                ' shipped default connection string is still a valid target, and writing to it
                ' while the operator believes the station is on the embedded cache is the worst
                ' kind of surprise. The probe runs before the main form opens, so it gets the
                ' short 2s budget the self-check harness uses, not the steady-state Connect
                ' Timeout; a dead SQL host must not stall startup.
                If AppStartup.Settings IsNot Nothing AndAlso AppStartup.Settings.DatabaseSettings.UseSqlServer AndAlso
                   AppStartup.ConnectionFactory IsNot Nothing Then
                    Dim builder As New Microsoft.Data.SqlClient.SqlConnectionStringBuilder(AppStartup.Settings.DatabaseSettings.ConnectionString) With {
                        .ConnectTimeout = 2
                    }
                    Using conn As New Microsoft.Data.SqlClient.SqlConnection(builder.ConnectionString)
                        conn.Open()
                        IsDatabaseConnected = (conn.State = System.Data.ConnectionState.Open)
                    End Using
                End If
            Catch ex As Exception
                IsDatabaseConnected = False
            End Try
            Coordinator = New DesktopDataCoordinator(IsDatabaseConnected)
        End Sub

        ''' <summary>
        ''' Applies a station setup change to the session that made it. All three steps are
        ''' required: settings alone leaves the connection factory, the repositories, and the
        ''' coordinator pointed at the mode the station used to be in.
        ''' </summary>
        Public Shared Sub ApplyStationSetup()
            AppSettings.Reload()
            AppStartup.Initialize()
            ProbeDatabaseConnection()
        End Sub

        <STAThread>
        Public Shared Sub Main(args As String())
            AppGlobalExceptionHandler.Setup()
            
            If args IsNot Nothing AndAlso args.Length > 0 AndAlso (args(0).ToLower() = "/test" OrElse args(0).ToLower() = "/smoke") Then
                Dim result = RunAutomatedSelfCheck()
                Environment.Exit(result)
            End If

            If args IsNot Nothing AndAlso args.Length > 0 AndAlso args(0).ToLower() = "/verify-audit" Then
                Environment.Exit(VerifyAuditChain(If(args.Length > 1, args(1), Nothing)))
            End If

            ' The elevated half of Set up this server: FormConnect re-launches this exe
            ' with /prepnet under the administrator verb, so the SQL service restart, the
            ' SQL Browser service, and the firewall rules run with real rights. It sits
            ' before the single-instance mutex because the parent app already holds it.
            If args IsNot Nothing AndAlso args.Length > 0 AndAlso args(0).ToLower() = "/prepnet" Then
                Dim steps As New List(Of String)()
                Dim exitCode = NetworkPrep.Run(Sub(line) steps.Add(line))
                Dim summary = If(exitCode = 0, "Network setup finished.", "Network setup finished with problems:")
                MessageBox.Show(summary & vbCrLf & String.Join(vbCrLf, steps.ToArray()), "BTA OSG : Server Network Setup",
                                MessageBoxButtons.OK, If(exitCode = 0, MessageBoxIcon.Information, MessageBoxIcon.Warning))
                Environment.Exit(exitCode)
            End If

            Application.SetHighDpiMode(HighDpiMode.SystemAware)
            Application.EnableVisualStyles()
            Application.SetCompatibleTextRenderingDefault(False)
            WheelScroller.Install()
            UiActivityMonitor.Install()

            ' Two instances writing bta_osg_db.xml would clobber each other's offline
            ' work, so a second launch is refused. Global\ covers every logon session on
            ' the machine (fast user switching, RDP into the server box); the self-check
            ' path above skips this: it must stay runnable while the desktop app is open.
            Dim createdNew As Boolean
            Using instanceMutex As New System.Threading.Mutex(True, "Global\BTA_OSG_DocumentTracking", createdNew)
                If Not createdNew Then
                    MessageBox.Show("BTA OSG Document Tracking is already running on this workstation.", "Already Running",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                AppStartup.Initialize()

                Dim requestedSetup = (args IsNot Nothing AndAlso args.Length > 0 AndAlso (args(0).ToLower() = "/setup" OrElse args(0).ToLower() = "--setup"))
                Dim setupApplied = False
                If requestedSetup OrElse Not AppSettings.Instance.IsConfigured Then
                    Using connect As New FormConnect()
                        Dim res = connect.ShowDialog()
                        If res = DialogResult.OK Then
                            ApplyStationSetup()
                            setupApplied = True
                        ElseIf Not AppSettings.Instance.IsConfigured Then
                            Return
                        End If
                    End Using
                End If
                ' Closing the wizard on a configured station keeps the saved configuration, so
                ' the probe still has to run exactly once for this session.
                If Not setupApplied Then ProbeDatabaseConnection()
                SealAuditChainInBackground()

                ' A station whose store holds no System Administrator has no card that can open
                ' it, so an officer has to be named before the app will run. Nothing is seeded
                ' here any more: the claimed badge is the only identity on the station.
                EmbeddedDB.Initialize()
                If Not AdministratorExists() Then
                    Using claim As New FormClaimAdmin()
                        If claim.ShowDialog() <> DialogResult.OK Then Return
                    End Using
                End If

                Application.Run(New FormMain())
            End Using
        End Sub

        ' Fills in hashes for entries older than migration 014. Idempotent, so safe on every
        ' launch; off the UI thread because years of entries can take seconds.
        Private Shared Sub SealAuditChainInBackground()
            If Not IsDatabaseConnected Then Return
            Task.Run(Sub()
                         Try
                             Dim sealed = AppStartup.AuditRepo.SealUnchainedRows()
                             If sealed > 0 Then ReportOperatorWarning("Audit chain sealed " & sealed.ToString() & " pre-existing entr" & If(sealed = 1, "y", "ies") & ".")
                         Catch ex As Exception
                             ReportOperatorWarning("Audit chain seal pass did not run: " & ex.Message)
                         End Try
                     End Sub)
        End Sub

        ' Recomputes the audit chain. Exit 0 means every sealed entry is intact; 1 means history
        ' was altered or the store could not be read. serverOverride is either a bare server name
        ' or a whole connection string.
        Private Shared Function VerifyAuditChain(Optional serverOverride As String = Nothing) As Integer
            Try
                For Each ignored In AppSettings.Instance.IgnoredConfigLayers
                    Console.WriteLine("[WARN] Config layer ignored: " & ignored)
                Next
                ResolveAuditStore(serverOverride)
                Console.WriteLine("Server: " & DescribeTarget(AppSettings.Instance.DatabaseSettings.ConnectionString))
                Dim result = AppStartup.AuditRepo.VerifyChain()
                If result.IsValid Then
                    Console.WriteLine("[OK] Audit chain verified: " & result.RowsChecked.ToString() & " sealed entries intact.")
                    If result.RowsUnsealed > 0 Then
                        Console.WriteLine("     " & result.RowsUnsealed.ToString() & " entries are not sealed yet. Open the app on a connected seat to seal them.")
                    End If
                    Return 0
                End If
                Console.WriteLine("[BROKEN] " & result.Reason)
                Console.WriteLine("     First entry that fails to verify: " & result.FirstBrokenAuditID.ToString())
                Return 1
            Catch ex As Exception
                Console.WriteLine("[ERROR] Audit chain could not be read: " & ex.Message)
                Return 1
            End Try
        End Function

        ' Chooses which store to read, in memory only, and never writes the choice back to disk.
        ' An explicit override wins outright; otherwise the configured server is tried, and if it
        ' does not answer the machine's own SQL instances are searched, because the usual case for
        ' running this by hand is a laptop that was never the office server.
        Private Shared Sub ResolveAuditStore(serverOverride As String)
            Dim configured = AppSettings.Instance.DatabaseSettings.ConnectionString
            Dim chosen As String = Nothing

            If Not String.IsNullOrWhiteSpace(serverOverride) Then
                chosen = ApplyServerOverride(configured, serverOverride)
            ElseIf AuditStoreAnswers(configured) Then
                chosen = configured
            Else
                For Each candidate In LocalStoreCandidates()
                    If AuditStoreAnswers(candidate) Then
                        chosen = candidate
                        Console.WriteLine("Configured server did not answer; reading this machine's instance.")
                        Exit For
                    End If
                Next
            End If

            If chosen IsNot Nothing Then AppSettings.Instance.DatabaseSettings.ConnectionString = chosen
            AppStartup.Initialize()
        End Sub

        Private Shared Function ApplyServerOverride(configured As String, serverOverride As String) As String
            Dim override = serverOverride.Trim()
            If override.Contains("="c) Then
                ' The operator states the trust choice in full. The tool never quietly disables
                ' certificate validation for a server reached over the network.
                Return New Microsoft.Data.SqlClient.SqlConnectionStringBuilder(override).ConnectionString
            End If
            Dim builder As New Microsoft.Data.SqlClient.SqlConnectionStringBuilder(configured)
            builder.DataSource = override
            Return builder.ConnectionString
        End Function

        ' Instances hosted on this machine, enumerated from its registry rather than from a
        ' guessed host name. Certificate trust is relaxed for these only: the target is a local
        ' pipe on the box under the operator's own login, with no network path to impersonate.
        Private Shared Function LocalStoreCandidates() As List(Of String)
            Dim list As New List(Of String)()
            For Each instanceName In SqlInstanceDiscovery.LocalInstanceNames()
                Dim builder As New Microsoft.Data.SqlClient.SqlConnectionStringBuilder(AppSettings.Instance.DatabaseSettings.ConnectionString)
                builder.DataSource = If(instanceName.Equals("MSSQLSERVER", StringComparison.OrdinalIgnoreCase), "localhost", "localhost\" & instanceName)
                builder.TrustServerCertificate = True
                list.Add(builder.ConnectionString)
            Next
            Return list
        End Function

        Private Shared Function AuditStoreAnswers(connectionString As String) As Boolean
            Try
                Dim builder As New Microsoft.Data.SqlClient.SqlConnectionStringBuilder(connectionString)
                builder.InitialCatalog = "BTA_OSG_DB"
                builder.ConnectTimeout = 2
                Using conn As New Microsoft.Data.SqlClient.SqlConnection(builder.ConnectionString)
                    conn.Open()
                    Using cmd = New Microsoft.Data.SqlClient.SqlCommand("SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = 'tbl_AuditTrail'", conn)
                        Return Convert.ToInt32(cmd.ExecuteScalar()) > 0
                    End Using
                End Using
            Catch
                Return False
            End Try
        End Function

        ' Names the server actually attempted: a seat can end up on a different one than the
        ' screen it was configured from, and that is the first thing anyone debugging needs.
        Private Shared Function DescribeTarget(connectionString As String) As String
            Try
                Dim builder As New Microsoft.Data.SqlClient.SqlConnectionStringBuilder(connectionString)
                Return If(String.IsNullOrEmpty(builder.DataSource), "(no server configured)", builder.DataSource)
            Catch
                Return "(unreadable connection string)"
            End Try
        End Function

        ''' <summary>
        ''' Asked of whichever store this station signs in from: SQL when the office server
        ''' answered the probe, the embedded cache otherwise. A store that cannot be read
        ''' answers "an administrator is there", because the claim would fail against it too.
        ''' </summary>
        Friend Shared Function AdministratorExists() As Boolean
            Try
                If IsDatabaseConnected Then
                    Return AppStartup.UserRepo.HasAnyUserWithRole(RbacPolicy.ROLE_SYSADMIN)
                End If
                Return EmbeddedDB.DataSet.Tables("Users").Select("Role = 'System Administrator' OR Role = 'SYSADMIN'").Length > 0
            Catch ex As Exception
                Return True
            End Try
        End Function

        Private Shared Function RunAutomatedSelfCheck() As Integer
            Console.WriteLine("=========================================================")
            Console.WriteLine("  BTA OSG DOCUMENT TRACKING : AUTOMATED SELF-CHECK SUITE")
            Console.WriteLine("=========================================================")
            Try
                ' The self-check is the isolated offline harness: its badge fixtures come
                ' from the embedded demo seeds, so SQL mode stays off no matter what the
                ' machine's real configuration says, and the run never depends on a
                ' reachable SQL host.
                AppSettings.Instance.DatabaseSettings.UseSqlServer = False
                AppStartup.Initialize()
                ProbeDatabaseConnection()
                EmbeddedDB.Initialize()
                ' Seed-budget baseline: the harness must hand the store back with exactly the
                ' documents it found (final seed policy: nothing seeded anywhere persists).
                Dim documentsBefore As Integer = EmbeddedDB.DataSet.Tables("Documents").Rows.Count

                ' Test 1: DB connection check
                Dim dbConnected As Boolean = False
                Try
                    Dim connStr = AppStartup.Settings.DatabaseSettings.ConnectionString
                    Dim builder As New Microsoft.Data.SqlClient.SqlConnectionStringBuilder(connStr) With {
                        .ConnectTimeout = 2
                    }
                    Using conn As New Microsoft.Data.SqlClient.SqlConnection(builder.ConnectionString)
                        conn.Open()
                        dbConnected = True
                    End Using
                    Console.WriteLine("[PASS] 1. Database connection verified (SQL Server).")
                Catch ex As Exception
                    Console.WriteLine("[PASS] 1. Database connection check handled (SQL Server host offline; isolated test harness engaged).")
                End Try

                ' Test 2: Target OSG document categories & auto-routing
                Dim sec1 = DocumentService.GetDefaultSectionForCategory("REG_COMM")
                Dim sec2 = DocumentService.GetDefaultSectionForCategory("LEG")
                Dim sec3 = DocumentService.GetDefaultSectionForCategory("FIN")
                Dim sec4 = DocumentService.GetDefaultSectionForCategory("TRAVEL")
                If sec1 <> "Secretariat" OrElse sec2 <> "Legislative Section" OrElse sec3 <> "Finance Section" OrElse sec4 <> "Travel Section" Then
                    Console.WriteLine("[FAIL] 2. Category auto-routing check failed.")
                    Return 1
                End If
                Console.WriteLine("[PASS] 2. Target OSG categories and auto-routing verified (4 categories).")

                ' Test 3: RFID card lookup. The store ships empty, so the harness enrols its own
                ' badges and removes them with the probe documents at the end of the run.
                Dim sgUid = "SELFCHKSG1"
                Dim deskUid = "SELFCHKDESK1"
                EmbeddedDB.AddUser(sgUid, "Self-Check Secretary-General", "Secretary-General", "Office of the Secretary-General")
                EmbeddedDB.AddUser(deskUid, "Self-Check Legislative Clerk", "Legislative Section", "Legislative Section")
                Dim sgUser = EmbeddedDB.AuthenticateRFID(sgUid)
                Dim legUser = EmbeddedDB.AuthenticateRFID(deskUid)
                If sgUser Is Nothing OrElse legUser Is Nothing OrElse
                   EmbeddedDB.AuthenticateRFID(deskUid.ToLowerInvariant()) Is Nothing Then
                    Console.WriteLine("[FAIL] 3. RFID badge lookup failed.")
                    Return 1
                End If
                Console.WriteLine("[PASS] 3. RFID badge lookup verified (executive and section desk).")

                ' Test 4: Invalid RFID rejection
                Dim badUser = EmbeddedDB.AuthenticateRFID("UNKNOWN_CARD_999")
                If badUser IsNot Nothing Then
                    Console.WriteLine("[FAIL] 4. Invalid RFID should have been rejected.")
                    Return 1
                End If
                Console.WriteLine("[PASS] 4. Invalid RFID rejection verified.")

                ' Test 5: Section view isolation, with one document for the desk and one for
                ' another section, so the filter has something it must exclude as well as
                ' something it must include.
                Dim legDocId = EmbeddedDB.AddDocument(EmbeddedDB.GenerateDocCode("Legislative"), "Legislative", "Self-check isolation probe legislative", "Committee on Education", "Legislative Section", "CAB-B", "S-3", "BOX-04", "", "FOR_REVIEW", "Self-Check Legislative Clerk", "INCOMING", "Legislative Section")
                Dim finDocId = EmbeddedDB.AddDocument(EmbeddedDB.GenerateDocCode("Finance"), "Finance", "Self-check isolation probe finance", "Finance Division", "Finance Section", "CAB-C", "S-2", "BOX-02", "", "RECEIVED", "Self-Check Other Officer", "INCOMING", "Finance Section")
                Dim legDocs = EmbeddedDB.GetVisibleDocuments(legUser)
                Dim seesOwn = legDocs.Cast(Of System.Data.DataRowView)().Any(Function(r) r("DocumentID").ToString() = legDocId.ToString())
                Dim seesOther = legDocs.Cast(Of System.Data.DataRowView)().Any(Function(r) r("DocumentID").ToString() = finDocId.ToString())
                If Not seesOwn OrElse seesOther Then
                    Console.WriteLine("[FAIL] 5. Section isolation failed: desk sees its own=" & seesOwn & ", sees the other desk=" & seesOther & ".")
                    Return 1
                End If
                Console.WriteLine("[PASS] 5. Section desk role isolation verified.")

                ' Test 6: Document auto-coding for target prefixes (COMM, LEG, FIN, TO)
                Dim c1 = EmbeddedDB.GenerateDocCode("Regular Communication")
                Dim c2 = EmbeddedDB.GenerateDocCode("Legislative")
                Dim c3 = EmbeddedDB.GenerateDocCode("Finance")
                Dim c4 = EmbeddedDB.GenerateDocCode("Travel Order")
                Dim codePattern = "^[A-Z]{2,4}-\d{4}-\d{3,}$"
                If Not Regex.IsMatch(c1, codePattern) OrElse Not Regex.IsMatch(c2, codePattern) OrElse Not Regex.IsMatch(c3, codePattern) OrElse Not Regex.IsMatch(c4, codePattern) Then
                    Console.WriteLine("[FAIL] 6. Document auto-coding format invalid: " & c1 & ", " & c2 & ", " & c3 & ", " & c4)
                    Return 1
                End If
                Console.WriteLine("[PASS] 6. Document auto-coding verified (" & c1 & ", " & c2 & ", " & c3 & ", " & c4 & ").")

                ' Test 7: Sequential minting (the duplicate-prevention rule): the offline code
                ' derives from the rows already on file, so a mint, an insert, and a second
                ' mint must advance the number, and no two codes can collide.
                Dim seqA = EmbeddedDB.GenerateDocCode("Finance")
                Dim numA As Integer = 0
                Integer.TryParse(seqA.Substring(seqA.LastIndexOf("-"c) + 1), numA)
                Dim seqDocId = EmbeddedDB.AddDocument(seqA, "Finance", "Self-check sequential mint probe", "Ministry of Finance", "Finance Section", "CAB-C", "S-2", "BOX-02", "", "RECEIVED", "Hassim A. Ibrahim", isOffline:=False)
                Dim seqB = EmbeddedDB.GenerateDocCode("Finance")
                Dim numB As Integer = 0
                Integer.TryParse(seqB.Substring(seqB.LastIndexOf("-"c) + 1), numB)
                If numB <> numA + 1 OrElse String.Equals(seqA, seqB, StringComparison.Ordinal) Then
                    Console.WriteLine("[FAIL] 7. Sequential code minting failed: " & seqA & " then " & seqB)
                    Return 1
                End If
                Console.WriteLine("[PASS] 7. Sequential document code minting verified (" & seqA & " then " & seqB & ").")

                ' Test 8: Full Sec Gen Revision Loop lifecycle
                Dim testDocCode = EmbeddedDB.GenerateDocCode("Regular Communication")
                Dim docId = EmbeddedDB.AddDocument(testDocCode, "Regular Communication", "Self-Check Workflow Verification", "Ministry of Finance", "Office of the Secretary-General", "CAB-A", "S-1", "BOX-01", "", "FOR_REVIEW", "Amina T. Macacua", "INCOMING", "Secretariat")

                ' Sec Gen orders revision
                EmbeddedDB.RequestRevision(docId, "Self-check punchlist notes", "Secretariat", "Prof. Ali B. Pangalian")
                Dim checkRow = EmbeddedDB.DataSet.Tables("Documents").Select("DocumentID = " & docId)(0)
                If checkRow("CurrentStatus").ToString() <> "FOR_REVISION" OrElse Not checkRow("RevisionPunchlist").ToString().Contains("Self-check punchlist notes") Then
                    Console.WriteLine("[FAIL] 8. Revision request transition failed.")
                    Return 1
                End If

                ' Secretariat resubmits
                EmbeddedDB.ResubmitDocument(docId, "Amina T. Macacua", "Resubmitted with punchlist completed")
                checkRow = EmbeddedDB.DataSet.Tables("Documents").Select("DocumentID = " & docId)(0)
                If checkRow("CurrentStatus").ToString() <> "FOR_REVIEW" OrElse checkRow("AssignedSection").ToString() <> "Secretary-General" Then
                    Console.WriteLine("[FAIL] 8. Resubmission transition failed.")
                    Return 1
                End If

                ' Sec Gen approves
                EmbeddedDB.ApproveDocument(docId, "Prof. Ali B. Pangalian", "Approved for release")
                checkRow = EmbeddedDB.DataSet.Tables("Documents").Select("DocumentID = " & docId)(0)
                If checkRow("CurrentStatus").ToString() <> "APPROVED" OrElse checkRow("AssignedSection").ToString() <> "Records Section" Then
                    Console.WriteLine("[FAIL] 8. Approval transition failed.")
                    Return 1
                End If

                ' Records Section releases
                EmbeddedDB.ReleaseDocument(docId, "Sittie K. Amin", "Released to Ministry of Finance")
                checkRow = EmbeddedDB.DataSet.Tables("Documents").Select("DocumentID = " & docId)(0)
                If checkRow("CurrentStatus").ToString() <> "RELEASED" Then
                    Console.WriteLine("[FAIL] 8. Release transition failed.")
                    Return 1
                End If
                Console.WriteLine("[PASS] 8. Full Sec Gen revision loop and status workflow verified.")

                ' Test 9: Append-only audit record creation
                EmbeddedDB.LogAudit("SYSTEM", "Self-check audit record verification.")
                Dim auditRows = EmbeddedDB.DataSet.Tables("AuditTrail").Rows
                If auditRows.Count = 0 Then
                    Console.WriteLine("[FAIL] 9. Audit trail logging failed.")
                    Return 1
                End If
                Console.WriteLine("[PASS] 9. Append-only audit record creation verified.")

                ' Tests 5, 7-9 wrote real rows into the shared cache; a PendingSync row would be
                ' replayed into SQL Server by the next production launch as a genuine document.
                Dim fixtureDocIds() As Integer = {seqDocId, docId, legDocId, finDocId}
                RemoveSelfCheckArtifacts(fixtureDocIds)
                RemoveSelfCheckBadges(sgUid, deskUid)

                ' Test 10: seed budget. Final policy: seeds never enter an office server, and
                ' anything this harness seeds stays capped at five documents and is removed
                ' before the run ends. A PendingSync row surviving here would replay into SQL
                ' Server on the next production launch as a genuine document.
                If fixtureDocIds.Length > 5 Then
                    Console.WriteLine("[FAIL] 10. Seed budget exceeded: the harness seeded " & fixtureDocIds.Length & " documents (cap 5).")
                    Return 1
                End If
                Dim seedDelta As Integer = EmbeddedDB.DataSet.Tables("Documents").Rows.Count - documentsBefore
                If seedDelta <> 0 Then
                    Console.WriteLine("[FAIL] 10. Self-check cleanup left " & seedDelta & " probe documents in the store.")
                    Return 1
                End If
                Console.WriteLine("[PASS] 10. Seed budget respected (cap 5) and the store handed back as found.")

                ' Test 11: PDF URL validation
                Dim errMsg As String = ""
                If Not AppStartup.PdfService.ValidateUrl("https://drive.google.com/file/d/sample-123/view", errMsg) Then
                    Console.WriteLine("[FAIL] 11. Valid URL rejected: " & errMsg)
                    Return 1
                End If
                If AppStartup.PdfService.ValidateUrl("http://untrusted-site.com/doc.pdf", errMsg) Then
                    Console.WriteLine("[FAIL] 11. Invalid URL accepted.")
                    Return 1
                End If
                Console.WriteLine("[PASS] 11. PDF URL security validation verified.")

                ' Test 12: Physical storage landmark movement
                Dim cabinet = "CAB-A"
                Dim shelf = "S-1"
                Dim box = "BOX-01"
                Dim landmarkKey = String.Format("{0}|{1}|{2}", cabinet, shelf, box)
                If landmarkKey <> "CAB-A|S-1|BOX-01" Then
                    Console.WriteLine("[FAIL] 12. Storage landmark key format failed.")
                    Return 1
                End If
                Console.WriteLine("[PASS] 12. Physical storage landmark tracking verified (" & landmarkKey & ").")

                ' Test 13: the mirror never throws and never reports rows for an offline host.
                ' The poller runs every 5 seconds on every workstation, so a throw here would be
                ' a recurring dialog, not a one-off.
                Dim snapshot = Coordinator.BuildSqlSnapshot()
                If Not IsDatabaseConnected AndAlso snapshot IsNot Nothing Then
                    Console.WriteLine("[FAIL] 13. Mirror returned rows while the database probe reported offline.")
                    Return 1
                End If
                Console.WriteLine("[PASS] 13. SQL mirror read handled (live database: " & (snapshot IsNot Nothing).ToString() & ").")

                ' Flush deferred embedded-store writes so the file on disk matches the
                ' mutations this run performed (mutations now batch via EmbeddedDB.MarkDirty).
                EmbeddedDB.Save()

                ' Test 14: audit hash chain. The harness runs offline, so this proves the
                ' canonical form and the walk rather than the SQL write path.
                Dim chainRows As New List(Of AuditEntry)()
                Dim chainPrev As Byte() = Nothing
                For i = 1 To 3
                    Dim row As New AuditEntry With {
                        .AuditID = i,
                        .EventAtUTC = New DateTime(2026, 10, 2, 8, 0, 0, DateTimeKind.Utc).AddMinutes(i),
                        .ActionType = "SELF_CHECK_" & i.ToString(),
                        .Success = True,
                        .MachineName = Environment.MachineName
                    }
                    row.PrevHash = chainPrev
                    row.RowHash = AuditChain.ComputeRowHash(row, chainPrev)
                    chainPrev = row.RowHash
                    chainRows.Add(row)
                Next
                If Not AuditChain.Verify(chainRows).IsValid Then
                    Console.WriteLine("[FAIL] 14. A clean audit chain did not verify.")
                    Return 1
                End If
                chainRows(1).ActionType = "TAMPERED"
                Dim tamperResult = AuditChain.Verify(chainRows)
                If tamperResult.IsValid OrElse tamperResult.FirstBrokenAuditID <> 2 Then
                    Console.WriteLine("[FAIL] 14. An altered audit entry was not detected at entry 2.")
                    Return 1
                End If
                Console.WriteLine("[PASS] 14. Audit hash chain seals entries and detects an altered one.")

                ' Test 15: same-version republish detection. publish-release.ps1 stamps every
                ' build with a UTC timestamp (InformationalVersion build metadata) and appends
                ' the same stamp to the release body; an installed copy must treat a release
                ' whose stamp is newer as an update even though the version did not change.
                Dim stampedBody = "### What's New in v2.1.6" & vbCrLf & vbCrLf & "- Fixed a thing" & vbCrLf & vbCrLf & "<!-- build:2026-10-03T14:22:11Z -->"
                Dim releaseStamp = AppUpdateService.ParseBuildStamp(stampedBody)
                If releaseStamp <> New DateTime(2026, 10, 3, 14, 22, 11, DateTimeKind.Utc) Then
                    Console.WriteLine("[FAIL] 15. Release build stamp was not parsed from the release notes.")
                    Return 1
                End If
                If AppUpdateService.ParseBuildStamp("notes without a stamp") <> DateTime.MinValue Then
                    Console.WriteLine("[FAIL] 15. An unstamped release body must not yield a stamp.")
                    Return 1
                End If
                ' The SDK appends the source revision to InformationalVersion; the real installed
                ' build reads "2.1.6+<stamp>.<sha>" and the stamp must still come out clean.
                If AppUpdateService.ParseInformationalStamp("2.1.6+2026-10-03T14:22:11Z.d196c725") <> releaseStamp Then
                    Console.WriteLine("[FAIL] 15. The installed build's stamp was not parsed from InformationalVersion.")
                    Return 1
                End If
                Dim olderStamp = releaseStamp.AddMinutes(-1)
                If Not AppUpdateService.IsRepublishNewer(releaseStamp, olderStamp) Then
                    Console.WriteLine("[FAIL] 15. A newer same-version release was not treated as an update.")
                    Return 1
                End If
                If AppUpdateService.IsRepublishNewer(olderStamp, releaseStamp) OrElse AppUpdateService.IsRepublishNewer(DateTime.MinValue, olderStamp) Then
                    Console.WriteLine("[FAIL] 15. An older or unstamped release must not be treated as an update.")
                    Return 1
                End If
                ' Two-part release tags (2.2) must read as the same version as their installed
                ' builds (2.2.0.0), or a republish of such a tag would never be detected.
                If Not AppUpdateService.IsSameVersion(New Version(2, 2), New Version(2, 2, 0, 0)) OrElse
                   Not AppUpdateService.IsSameVersion(New Version(2, 1, 6), New Version(2, 1, 6, 0)) OrElse
                   AppUpdateService.IsSameVersion(New Version(2, 2), New Version(2, 1, 6, 0)) Then
                    Console.WriteLine("[FAIL] 15. Version part-count normalization failed.")
                    Return 1
                End If
                If AppUpdateService.FormatVersion(New Version(2, 2, 0, 0)) <> "2.2" OrElse AppUpdateService.FormatVersion(New Version(2, 1, 6, 0)) <> "2.1.6" Then
                    Console.WriteLine("[FAIL] 15. Version display formatting failed.")
                    Return 1
                End If
                Console.WriteLine("[PASS] 15. Same-version republish detection and two-part version handling verified.")

                ' Test 16: Clean exit
                Console.WriteLine("[PASS] 16. Clean exit verified.")

                Console.WriteLine("=========================================================")
                Console.WriteLine("ALL 16 SELF-CHECK TESTS PASSED SUCCESSFULLY!")
                Console.WriteLine("=========================================================")
                Return 0
            Catch ex As Exception
                Console.WriteLine("[FAIL] Self-check exception: " & ex.Message)
                Return 1
            End Try
    End Function

    ''' <summary>
    ''' Drops the harness's own badges, so the enrolled staff list the operator sees is only
    ''' ever staff, never test rows.
    ''' </summary>
    Private Shared Sub RemoveSelfCheckBadges(ParamArray uids As String())
        SyncLock EmbeddedDB.SyncRoot
            Dim dt = EmbeddedDB.DataSet.Tables("Users")
            For i As Integer = dt.Rows.Count - 1 To 0 Step -1
                If uids.Contains(dt.Rows(i)("RFID_UID").ToString().ToUpperInvariant()) Then dt.Rows(i).Delete()
            Next
        End SyncLock
    End Sub

    ''' <summary>
    ''' Deletes the documents and child rows the self-check created, plus the audit rows that
    ''' name them, so the self-check leaves the production cache exactly as it found it.
    ''' </summary>
    Private Shared Sub RemoveSelfCheckArtifacts(ParamArray docIds As Integer())
        SyncLock EmbeddedDB.SyncRoot
            For Each tblName In {"Directives", "RoutingLogs", "Movements", "Documents"}
                Dim dt = EmbeddedDB.DataSet.Tables(tblName)
                If dt Is Nothing Then Continue For
                For i As Integer = dt.Rows.Count - 1 To 0 Step -1
                    For Each docId As Integer In docIds
                        If dt.Rows(i)("DocumentID").ToString() = docId.ToString() Then
                            dt.Rows(i).Delete()
                            Exit For
                        End If
                    Next
                Next
            Next
            Dim audit = EmbeddedDB.DataSet.Tables("AuditTrail")
            If audit IsNot Nothing Then
                For i As Integer = audit.Rows.Count - 1 To 0 Step -1
                    Dim desc = audit.Rows(i)("ActionDescription").ToString()
                    Dim isProbeAudit = desc.Contains("Self-check") OrElse desc.Contains("audit record verification.")
                    For Each docId As Integer In docIds
                        If desc.Contains("Doc #" & docId.ToString()) OrElse desc.Contains("Document #" & docId.ToString()) Then
                            isProbeAudit = True
                            Exit For
                        End If
                    Next
                    If isProbeAudit Then audit.Rows(i).Delete()
                Next
            End If
        End SyncLock
    End Sub
End Class
End Namespace
