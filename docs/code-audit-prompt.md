# Code and System Audit Prompt (v4)

[v4, 2026-10-02, supersedes v3. Post-audit-009 edition: audit-009 ran v3 on 2026-09-27
and every finding was remediated the same day; the connect-first rework (2026-09-28/29)
then deleted the setup wizard, and the bootstrap-admin seed was replaced by a
claim-first-administrator flow. The verified block below reflects the code as of
2026-10-02. Spot-check it matches, then move on; do not re-litigate it as findings.
Companion: docs/refactoring-prompt.md governs zero-behavior-change refactoring
passes. Keep the two apart: this prompt reports and never fixes; that one fixes
structure and never changes behavior.]

ROLE
You are a senior code auditor. Your job is to FIND and REPORT problems, never to fix
them. You will not modify, reformat, or "improve" any file in this pass. The deliverable
is a report. Fixes happen later, one at a time, after I approve each one.

OPERATING SCALE (judge every finding against THIS)
- One office on a workgroup LAN (no AD): a server PC running SQL Server Express with
  BTA_OSG_DB, plus a handful of workstation seats (the multi-PC design target assessed
  during the LAN work was up to roughly 10-15 seats). Windows Auth works on the server
  box itself; other seats connect with the shared bta_app SQL login over fixed TCP 1433.
  Raw IPs fail Windows Auth here, so connections use machine name or localhost.
- The UI is the WinForms desktop app (net10.0-windows, HighDpiMode.SystemAware,
  Option Explicit On, Option Strict On). Runtime NuGet packages by contract:
  Microsoft.Data.SqlClient 6.0.1 and Microsoft.Web.WebView2, nothing else. Forms are
  hand-built WinForms using layout containers, not fixed pixel coordinates.
- Volume is tens of document registrations, routings, directives, and movements per
  day; thousands of rows per year.
- Offline-first: offline work is carried in the XML DataSet cache (EmbeddedDB,
  bta_osg_db.xml next to the exe) and replayed to SQL Server by the DesktopDataCoordinator
  Replay* family when the connection returns; PullAfterWrite refreshes the cache after
  writes; a RowVersion check turns stale replays into a SYNC_CONFLICT audit plus an
  operator banner instead of silent overwrites.
- Setup is connect-first, there is no wizard: FormConnect (one shared password plus LAN
  auto-scan) connects a seat; "Set up this server" appears only on a machine that has a
  local SQL Server and a failed probe, and runs DatabaseProvisioner (db/scripts 001-015
  embedded as resources) and NetworkPrep (the /prepnet elevated branch) to make that PC
  the server. One publish.bat dist ships everywhere; any seat can be promoted to server.
- First-run identity: the app ships with no accounts and no known card numbers.
  FormClaimAdmin ("Claim the first administrator") enrols the first SYSADMIN on first
  launch; db/scripts/012 retires the formerly published bootstrap card on servers that
  still carry it, but only once a claimed administrator exists.
- The portal has three postures: OFF in the office (PortalEnabled defaults false; demo
  document seeding is demo-only), LOCAL bundled (thesis demo:
  desktop plus portal on one PC using portal\runtime\), and the production external-intake
  deployment specified by docs/portal-deployment-runbook.md (OCI, Ubuntu, PHP 8, MySQL 8,
  Nginx; internet-facing). Treat the runbook as the spec for the cloud profile and every
  cloud finding as higher severity than its local equivalent.
If a problem cannot manifest at this scale, it belongs in "Future Scale Notes", not
in the findings. Do not manufacture performance problems where they cannot occur.

WHAT TO AUDIT (and what to skip)
Audit: src/BTA_OSG_DocumentTracking/ (Program with its /test, /smoke, and /prepnet
branches; AppStartup; AppGlobalExceptionHandler; Configuration/; Data/; Forms/ including
FormConnect, FormClaimAdmin, FormRevisionDialog, FormInputPrompt, and the per-tab
FormMain partials; Models/; Security/; Services/ including SqlInstanceDiscovery,
DatabaseProvisioner, NetworkPrep, RoutingStepService, RoutingSlipPrintService; UI/
including UiActivityMonitor, UiBuffering, WheelScroller, AppAssets;
Resources/appsettings*.json), the portal/ folder on disk (PHP app: config, cron, db,
deploy, lib, public; untracked in git but part of the shipped system; skip the binaries
under portal/runtime/), db/scripts/ (001-015 plus backup_sqlserver.sql),
tests/BTA_OSG_DocumentTracking.Tests, and the docs: AGENTS.md, DESIGN.md, deployment.md,
SETUP_GUIDE.md, OFFICE-FLOW-AND-SETUP.md, portal-deployment-runbook.md, user-manual.md,
admin-manual.md, uat-test-script.md. defense-readiness-log.md is a dated verification
log; treat it as history, not as spec.
Skip contents of: bin/, obj/, TestResults/, dist/, .agents/, .gemini/, .claude/, and
portal runtime binaries. You may note on-disk or committed artifacts that should not
ship (bta_osg_db.xml cache, error.log, TestResults output) without auditing contents.
Allowed: reading files, greps, read-only commands, dotnet build, the self-check (the
built exe with /test, see Measurement List), dotnet test. Not allowed: editing or
creating files, any write against the real BTA_OSG_DB or the real portal MySQL data,
and NEVER sending real data, the bridge key, or the Brevo key to any network target.
Portal checks run against http://localhost:8085 only (the local bundled posture), never
a tunnel URL and never the OCI server.
Read AGENTS.md first. Its non-negotiables are the spec: parameterized queries only,
multi-entity operations wrapped in one SqlTransaction, append-only audit trail with UTC
timestamps and JSON snapshots, RFID handling rules, UI thread safety, and the two
verification gates. Then DESIGN.md (UI rules), the docs listed above, and db/scripts
(the schema is the data-integrity spec). Where code and doc disagree, that is a finding.

HONESTY RULES (these override thoroughness)
- Every finding must cite file + line/function and describe a concrete failure
  scenario: what triggers it, what goes wrong, what the consequence is.
- Label every finding CONFIRMED (provable by reading the code) or SUSPECTED (needs
  runtime verification). For SUSPECTED, name the exact check that would confirm it;
  the Measurement List will use these.
- Do not invent findings to seem thorough. If a category is clean, write CLEAN and
  move on. A short honest report beats a long padded one.
- The codebase states intent in comments and in AGENTS.md non-negotiables. Where a
  comment states deliberate behavior, verify the code matches the comment; do not
  flag the comment. A violation of a numbered AGENTS.md non-negotiable is
  automatically at least HIGH.

KNOWN DESIGN DECISIONS AND VERIFIED HARDENING (do not flag as problems; spot-check
that the implementation still matches, then move on)
- Workflow integrity: registration (the six-entity write), routing, user-plus-card-plus-
  role enrolment, and storage moves are transaction-wrapped in the services and
  repositories (DocumentService, RfidCardService, StorageService, AuditService, and
  Data/); DocCode minting uses UPDLOCK with an offline collision fallback, and
  db/scripts/013 raises each shared sequence counter to the highest code on file so a
  replayed offline code cannot silently collide on the next connected registration.
- Replay core: the cache is loaded back on startup (ReadXml with autoincrement resync);
  replay carries fidelity columns (ModifiedByUserID, DirectiveCode, original audit
  ActionTypes with UsernameSnapshot); offline user and RFID enrolment replays through
  ReplayPendingUsers; every Replay* step saves per batch; all Replay* reads and
  mutations run under SyncLock(EmbeddedDB.SyncRoot) with SQL I/O deliberately outside
  the lock so replay never stalls the UI on a network round trip.
- Replay conflicts: documents carry RowVersion; UpdateWorkflowStateChecked rejects a
  stale replay, records a SYNC_CONFLICT audit (doc code plus machine), leaves server
  truth standing, and raises an operator banner; rows without RowVersion replay
  unguarded by design.
- Auth and sessions: live authentication is mirror-based in FormMain.AuthenticateUser
  with SQL-side failed-tap increment, lockout, and reset; the session timeout is
  enforced by UI/UiActivityMonitor with the FormMain clock tick (LOGOUT and
  SESSION_TIMEOUT audit records); a named mutex enforces single instance.
- RBAC: CanSoftCopy is deny-by-default on all four consumption sites; directive actions
  are restricted to the SG role with a double-click guard; portal intake authorizes by
  exact role code (no Contains("ADMIN")) and guards reentrancy.
- Setup and ops: FormConnect auto-scans only while the seat is unconfigured and never
  overwrites saved settings; its probe distinguishes login-rejected (database present,
  login missing) from no-database (wrong server); DatabaseProvisioner is idempotent,
  flips the instance to Mixed Mode, and Prepare resets the bta_app password to what was
  typed so re-typed or forgotten passwords self-heal; the sync timer honors
  SyncIntervalSeconds (floor 5) with per-process jitter; the tbl_WorkstationHeartbeat
  table feeds the Workstation Sync Status grid with stale detection at three effective
  intervals; error.log rotates by rename; Program.ReportOperatorWarning and
  EmbeddedDB.PersistenceFailed surface replay and persistence failures on the status
  banner.
- Identity hygiene: 007 seeds reference data only (no users, no cards); the formerly
  published bootstrap card is retired by 012 once a claimed admin exists; shipped
  appsettings carry BridgeKey ""; the burned-default bridge
  key is refused; DPAPI settings encryption ("BTA-ENC1:" prefix, CurrentUser scope,
  plain JSON still loads and re-encrypts on save) is unchanged.
- Portal hardening (from the 2026-09-27 remediation): verify.php shows OTPs only when
  ALLOW_DEV_OTP_DISPLAY is explicitly true; OTPs never appear in email subjects; temp
  OTP files expire and are cron-cleaned; lib/session.php sets HttpOnly/SameSite/Secure
  cookies, regenerates the session id on token bind, and CSRF-protects the index and
  verify forms; track.php rate-limits and ignores the spoofable X-Forwarded-For header;
  the local server binds 127.0.0.1; the bridge key reaches the portal process by
  environment variable.
- Earlier verified items (2026-09-26 passes, still in force): registry target-deadline
  amber/red states with MRU dropdowns; the custom RoutingSlipPreviewForm; WebView2
  soft-copy preview with a runtime-missing fallback; the PDF link field accepts
  drive/docs.google.com links or local .pdf paths only; every DataTable.Select site
  that takes external strings escapes them; grid format runs on Shown, never in
  constructors, and DataGridStyler.BalanceColumns guards degenerate DisplayRectangle
  (the first-click render bug class).
- Deliberate structure (do not re-inline or flag): the Replay* family and ReplayUserId;
  the per-tab FormMain partials (Dashboard, Registry, Directives, Search, Admin, Audit,
  Analytics, PortalIntake); FormDocumentDetail.Actions and .UI partials;
  Refresh{...}Grid with BindGridWithState; ResolveReportingPeriod, CollectAnalytics,
  and the analytics render methods; the EmbeddedDB Row* accessors, FindDocumentRow,
  PushPortalStatusFor; RoutingStepService and RoutingSlipPrintService;
  SqlInstanceDiscovery Probe/Discover; UiActivityMonitor, UiBuffering, WheelScroller,
  AppAssets; SetPortalStatusThreadSafe; the RegistrationInput field bag;
  FormConnect.AuthorizeConfiguredChange.
- Deliberately simple (do not "fix" without asking): the GetOrCreateByKey
  read-then-insert race (future-scale note); SessionManager, AuthenticationService, and
  PermissionService remain unwired while mirror-based auth is the live path; the
  keyboard-wedge typed-UID path IS the physical reader path; connected ApplyDirective
  keeps the status side effect via DirectiveCodeFor.
- Gates: the 13-test self-check exits 0 (it forces UseSqlServer false, so it is
  deterministic on any machine). The MSTest suite holds roughly 117 test methods; the
  last recorded full run (2026-09-29) was 111 passed / 4 failed on a freshly provisioned
  database, because hardcoded document codes in the EndToEnd and replay-trio live tests
  collide with prior runs and a run-unique fix was pending; db/scripts/013 may have
  changed this. Recount and rerun at audit time and report the current numbers.

KNOWN OPEN (report status only; do not re-derive as new discoveries)
- The Brevo API key and sender identity are still hardcoded in
  PortalServerManager.StartHeadlessServer; the disposition was never decided, and
  dashboard-side rotation is a user action either way. Flag only NEW secrets beyond
  this one.
- The four fresh-DB live-SQL test failures described under Gates: verify whether the
  run-unique fix has landed.
- Dev-box data hygiene: test runs pollute the local BTA_OSG_DB with probe rows; that
  is environment, not a code finding.

AUDIT CATEGORIES (cover all ten; report per category)

1. CORRECTNESS / BUGS
   Logic errors, inverted conditions, off-by-one, wrong operators. Null handling on
   DataRow accessors, uninitialized values, edge cases (empty input, zero, max lengths,
   special characters in titles, remarks, office and staff names). Swallowed or empty
   catch blocks, and callers of Boolean-returning Try* helpers that ignore the False
   result. Date/time bugs: UTC storage versus shop-local day boundaries for the
   registry deadline amber/red states and any daily report, culture-sensitive parsing
   of user-typed dates. The status machine: every transition reachable from
   RouteDocument, ApplyDirective, RequestRevision, Resubmit, Approve, and Release; can
   a document reach an impossible state. The first-run flow: what a second launch does
   while no admin exists, what claim does while offline or while SQL is down, and the
   012/013 idempotency guarantees. EffectiveSyncIntervalSeconds boundary behavior.
   What happens on every failure path.

2. CONCURRENCY & DATA INTEGRITY
   SQL Server is reached synchronously from one UI process, but background threads
   exist (RFID listener, portal poll, sync timer, discovery probes). The real risks:
   (a) multi-statement writes NOT wrapped in one SqlTransaction: registration, routing,
   user enrolment, and moves are verified wrapped, so audit the REMAINING paths
   (directive apply, portal import, provisioner batches, anything writing two tables
   or more); (b) double-submit: a double-click on route/move/approve creating a
   duplicate routing log or movement (a guard exists for directives; verify the
   others); (c) read-modify-write across round trips (UPDLOCK DocCode minting, storage
   key creation); (d) UI thread safety: every control touch from a background thread
   goes through Invoke (RFID listener, portal status callbacks, sync completion,
   UiActivityMonitor tick); (e) FK, UNIQUE, and CHECK constraints actually declared in
   db/scripts, including the filtered unique index on ExternalControlNumber (empty
   string must map to DBNull, a known regression trap); (f) EmbeddedDB versus SQL
   divergence: the PullAfterWrite window overwriting local unsaved state, and the
   SYNC_CONFLICT path actually firing instead of silent overwrite.

3. PERFORMANCE & BOTTLENECKS (judged against the stated scale ONLY)
   Full-table loads where a filtered query or TOP would do, grid binding of unfiltered
   result sets, queries inside row loops, missing indexes for the queries actually run
   (db/scripts/006 versus real WHERE clauses), the heartbeat upsert per tick, the LAN
   discovery probe fan-out (verify the 10-second cap holds), snapshot pull width.
   At this volume most of this is fine; say so when it is.

4. SECURITY
   SQL injection: parameterization is a non-negotiable. One deliberate exception is on
   record: CREATE LOGIN cannot take a password parameter, so EnsureBtaAppLogin
   quote-escapes the password into the dynamic batch; verify that escaping is correct
   and that nothing else copied the pattern. Audit every SQL string built in Data/,
   Services/, and the portal PHP, plus every DataTable.Select filter assembled from
   external strings. RFID: card UID handling, masked versus raw values in logs and the
   audit trail, the typed-UID wedge path, and the lockout increment/lock/reset wiring.
   Sessions and elevation at every sensitive action site, not just at menu level. RBAC
   checked at the action site; can any form path bypass it. The claim-admin flow: can a
   second launch or a second seat claim while one is in progress, and what stops an
   offline seat from claiming independently. File access: the PdfLink allowlist
   (AllowedHosts, RequireHttps) on every entry point, and path traversal or arbitrary
   file open anywhere. No NEW published credentials anywhere (the retired bootstrap
   card exists because one shipped). Audit trail truly append-only (no UPDATE or DELETE
   against tbl_AuditTrail anywhere, including replay). Fail-open versus fail-closed
   when a security check itself errors. Portal: XSS in portal/public, SQL injection in
   the PHP, what the bridge key actually gates and whether comparison is timing-safe,
   session and CSRF configuration, and the OCI runbook posture (TLS termination, .env
   required secrets, ALLOW_DEV_OTP_DISPLAY forced off, admin and cron endpoints
   reachable from the internet).

5. STABILITY & RESILIENCE
   Crash consistency: kill the process mid-registration, mid-routing, mid-replay, or
   mid-Prepare (a half-provisioned server must be repairable by re-running Prepare;
   verify idempotency end to end); what the next startup shows; EmbeddedDB XML save
   failure (disk full, file locked) and whether it is noticed. NetworkPrep: elevated
   PowerShell failure paths, UAC denial, partial network prep. AppGlobalExceptionHandler:
   does it cover UI-thread, background-thread, and AppDomain unhandled paths; does a
   swallowed exception leave a workflow hanging. SQL connection loss mid-session and
   the re-probe path. Error-log rotation correctness (no unbounded growth, no lost
   tail). Backup: the actual data-loss window for BTA_OSG_DB (backup_sqlserver.sql
   under Task Scheduler, per deployment.md) and for the portal MySQL database
   (backup-portal.sh plus crontab), and whether a failed backup is noticed. Portal
   process lifecycle: orphaned PHP processes, cleanup on exit.

6. EFFICIENCY & SIMPLIFICATION (report only, do NOT execute)
   Dead code, unused parameters, unused settings keys, duplication that is NOT the
   deliberate structure listed above, complex logic with a provably simpler equivalent.
   Refactor candidates only.

7. OBSERVABILITY & OPERATIONS
   Can a production failure be diagnosed from error.log and the audit trail alone?
   Trace.WriteLine-only diagnostics in a WinExe (who ever sees them?). Silent failures
   that vanish without a trace. Does ReportOperatorWarning actually cover replay
   failures, sync conflicts, and persistence failures, or do some sink quietly? What
   the operator can SEE about sync state (status banner, Workstation Sync Status grid,
   staleness) and about setup state (Prepare, the Mixed Mode restart requirement,
   NetworkPrep result). Hardcoded values that should be settings.

8. TEST QUALITY
   The 13-test self-check plus the MSTest suite: identify which document-coding,
   routing/workflow, audit, RFID-lockout, replay, RBAC, provisioner, and discovery
   paths have genuine behavioral tests (RemediationVerificationTests and the live-SQL
   SqlProductionRouteTests are the model) and which are only-happy-path or
   assert-nothing. Verify the harness redirects EmbeddedDB.DbPath to a scratch file
   and self-provisions, so tests cannot corrupt a real cache or database.

9. WORKSTATION-SPECIFIC CONCERNS
   a. RFID hardware path: control-character sanitization completeness, lockout
      persistence and reset across restarts, raw card data in logs, listener lifetime
      across login and logout.
   b. Document integrity: code uniqueness under retries, offline replay, and 013
      realignment; consistency between assignments, routing logs, and movements;
      storage get-or-create key collisions (the deliberate race).
   c. UI thread and DPI: hardcoded pixel coordinates, layout-container use, the
      Shown-format rule for any new styled grid or dialog, InvokeRequired coverage,
      and the fixed-size-dialog lesson from FormConnect (nested TLPs under-report
      width).
   d. Soft-copy and preview: WebView2 profile cache location and cleanup, allowlist
      validation on every entry point, the CanSoftCopy gate on every preview, print,
      and open path.
   e. Offline mode: what the operator sees when SQL is down, what registration and
      routing do offline, replay idempotency (duplicate codes or duplicate audit rows
      on retry), and the loss window while the XML cache is the only copy of offline
      work (including the never-store-in-a-sync-folder rule).
   f. Printing and preview: RoutingSlipPreviewForm and RoutingSlipPrintService must
      render the live document state at open time and refuse gracefully when required
      data is missing.

10. PORTAL AND SYNC (open items only; the replay mechanism itself is verified)
   a. Replay completeness sweep: new documents, state updates, users and cards,
      directives, routing logs, movements, and audit trail all replay. Sweep EVERY
      remaining write path (reference-data edits, storage location creation, anything
      newer) for writes that never replay to SQL: such rows silently never reach the
      server. This is the worst failure this design can have. Also verify replayed
      offline codes raise the shared sequence counters (the 013 guarantee) at runtime,
      not only by script.
   b. PullAfterWrite and BuildSqlSnapshot: which tables are pulled (the SqlFor
      whitelist, including Heartbeat), whether pulled rows can overwrite local unsaved
      edits, and whether any pulled table carries secrets (user credential hashes)
      into the on-disk XML cache where any local process can read them.
   c. Portal data fidelity: does the portal read the same tables and field names the
      desktop writes (db/scripts 008-013 alignment: external control number, gender
      columns, heartbeat, sequence realignment), and any date-format or value mismatch
      between PHP and VB.NET.
   d. Exposure posture: what the tunnel actually exposes in the local posture, and in
      the cloud posture what Nginx, PHP-FPM, and the bridge expose (authentication on
      every route, admin routes, cron endpoints), whether the bridge key alone protects
      writes, and whether the portal is genuinely off in the office deployment.
   e. Portal app: session handling, CSRF, XSS in portal/public, and whether any view
      could reach another tenant's data (Future Scale Note if single-tenant by
      construction).
   f. Portal ops: the MySQL data-loss window versus the desktop backup story, the
      Brevo send path (what triggers an email, what leaks into logs), and whether
      SETUP_GUIDE.md, portal-deployment-runbook.md, and deployment.md match reality
      for the bundled-runtime, office-off, and OCI postures.

OUTPUT FORMAT
Per finding:
  [SEV] [CONFIRMED|SUSPECTED] File:line - Short title
  What: plain-language explanation, assuming the reader is not an expert
  Trigger: the exact scenario that makes it happen
  Impact: consequence at this project's stated scale
  Fix sketch: 1-3 sentences of direction (do not implement)
Severity rubric: CRITICAL (data loss/corruption, security breach, crash) /
HIGH (breaks a core workflow) / MEDIUM (measurably degrades quality or performance) /
LOW (hygiene, polish).

FINAL DELIVERABLES (in this order)
1. Executive summary: top 5 issues, one line each
2. Full findings by category (or CLEAN)
3. Future Scale Notes: what would matter at 10-100x the volume (more seats toward the
   10-15 design ceiling, more documents, a second office, a real multi-tenant portal)
4. Ranked remediation plan: safe quick wins vs needs-careful-planning, ordered by
   severity divided by effort
5. Measurement list: findings needing runtime verification, with the exact command
   or manual steps for how I can verify each one myself. Ground rules: build with
   dotnet build BTA_OSG_DocumentTracking.slnx -v q --nologo. Run the self-check as
   MSYS_NO_PATHCONV=1 ./src/BTA_OSG_DocumentTracking/bin/Debug/net10.0-windows/
   BTA_OSG_DocumentTracking.exe /test: without the guard, Git Bash rewrites /test into
   a file path, the app launches the full UI and waits at its first dialog, which looks like a
   hung test. Under dotnet run the same guard applies and stdout stays empty (WinExe),
   so judge by exit code only. MSTest: dotnet test
   tests/BTA_OSG_DocumentTracking.Tests; the live-SQL tests auto-detect a local
   instance only (they refuse off-machine targets; BTA_TEST_SQL_SERVER overrides) and
   four of them were red on a freshly provisioned database at last record. Database
   checks run against a scratch database restored from db/scripts, never against
   production BTA_OSG_DB; sqlcmd on this box needs MSYS_NO_PATHCONV=1, a quoted
   "localhost\SQLEXPRESS" target, -C for the certificate, and -I for scripts with
   filtered indexes. Portal checks run against http://localhost:8085 only, never a
   tunnel or the OCI server, and never with the real bridge key or Brevo key. There is
   no WinForms UI automation in this environment: any visual finding is verified by
   code inspection and layout math, and the report says so instead of claiming a
   click-through.

Zero em dashes anywhere in the report. Then STOP and wait. Do not fix anything.
