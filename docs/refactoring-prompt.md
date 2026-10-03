# Refactoring Pass Prompt (v2)

[v2, 2026-10-02, aligned to THIS repository: BTA OSG DocumentTracking, VB.NET
WinForms on .NET 10 with SQL Server. Supersedes the generic version; its
PROJECT-SPECIFIC INVARIANTS placeholder is now filled. One pass already ran here on
2026-09-26 (9 steps, net +83 lines), and the structure it created is listed under
PROTECTED STRUCTURE below: undoing that structure is the main failure mode of a
second pass. Refactoring and auditing are separate passes; docs/code-audit-prompt.md
governs find-and-report audits. Never fold audit findings into this pass, and never
fix anything inside this pass.]

MISSION
Improve internal code quality with ZERO behavior change. In scope:
- CHOPPING for readability: Extract Function/Method (long functions into small
  named ones) and Extract Class/Module (big multi-responsibility files split
  into cohesive pieces). These are the primary techniques.
- CLEANUP: rename toward domain language, remove duplication, inline useless
  indirection, delete dead code.
In-scope code: src/BTA_OSG_DocumentTracking/ and
tests/BTA_OSG_DocumentTracking.Tests/ (VB.NET).
OUT of scope: portal/ (PHP has no test harness on this machine, so a change there
cannot satisfy the baseline ritual; verify-by-inspection only, propose separately),
db/scripts/ (schema is behavior), Resources/appsettings*.json (settings shape is
behavior), vbproj wiring except where an approved step moves files and must update
EmbeddedResource LogicalNames. Bug fixes, features, behavior tweaks, dependency
upgrades, styling/format-only churn. If you find a bug, STOP and report it
separately, never fix it inside this pass.

BASELINE RITUAL (first, no exceptions)
1. The gates are fixed for this project; do not re-detect them:
   - Build: dotnet build BTA_OSG_DocumentTracking.slnx -v q --nologo (0 errors)
   - Self-check: MSYS_NO_PATHCONV=1 ./src/BTA_OSG_DocumentTracking/bin/
     Debug/net10.0-windows/BTA_OSG_DocumentTracking.exe /test (exit 0, 13/13).
     Without the env guard, Git Bash rewrites /test into a file path and the full
     UI launches. Under dotnet run, stdout stays empty (WinExe), so the exit code
     is the only signal.
   - Suite: dotnet test tests/BTA_OSG_DocumentTracking.Tests
2. Run all three. Record exact results INCLUDING failures. The suite may have
   pre-existing reds on a freshly provisioned database (four live-SQL tests with
   hardcoded document codes were red at last record). IDENTICAL before and after
   means the same set of pass and fail, not all-green. List pre-existing reds in
   the baseline; their remediation is bug fixing and is out of scope here.
3. The suite self-provisions: it redirects EmbeddedDB.DbPath to a scratch file and
   provisions a local SQL Server instance only (never off-machine). If the suite
   cannot run on this machine at all, stop and report instead of proceeding.
UI-visible behavior has no automated check in this environment (no WinForms UI
automation): proofs for UI-facing steps are code inspection plus layout math,
stated as such. Never claim a click-through that did not happen.

RULES OF ENGAGEMENT
- One refactoring at a time. Run all three gates. Only then continue. Do not
  commit unless the user asks.
- Split by RESPONSIBILITY, never by line count. New seams follow the existing
  ones (per-tab FormMain partials, FormDocumentDetail .Actions/.UI, Services,
  Data repositories). If total line count goes UP after an extract, that is fine
  and expected; readability is the metric, not brevity.
- Thresholds: act on functions over ~40 lines or with more than one nesting
  level; act on logic duplicated in 2+ places. Layout-builder methods
  (InitializeUI, InitializeForm, and control-tree blocks) are conventionally long
  and mechanical: do not split them for line count; extract only behavioral code
  buried in layout, such as event wiring or state setup.
- Honesty clause: a previous protocol pass already restructured this code
  deliberately. If a file is already readable and cohesive, write CLEAN and move
  on. Manufactured work is the failure mode here, not missed refactors.
- No new frameworks, libraries, or architectural layers. The NuGet contract is
  Microsoft.Data.SqlClient and WebView2, nothing else. Code lives in its existing
  layer (Forms, then Services, then Data repositories and Models, plus Security/,
  UI/, Configuration/). If a helper exists, move code INTO it; never create a
  parallel structure.
- Follow the project's conventions: Option Explicit On and Option Strict On
  (explicit conversions, no late binding); CivicCalmTheme and AppAssets for all
  styling, no ad-hoc colors or fonts; comments state constraints the code cannot
  show and never narrate changes (antislop-code); no em dashes anywhere.
- PROTECTED STRUCTURE (created deliberately by the 2026-09-26 pass; do NOT
  re-inline, merge away, or flag as duplication): the Replay* family and
  ReplayUserId in DesktopDataCoordinator; the per-tab FormMain partials (Dashboard,
  Registry, Directives, Search, Admin, Audit, Analytics, PortalIntake);
  Refresh{...}Grid with BindGridWithState; ResolveReportingPeriod, CollectAnalytics,
  and the analytics render methods; the EmbeddedDB Row* accessors, FindDocumentRow,
  PushPortalStatusFor; RoutingStepService and RoutingSlipPrintService;
  SqlInstanceDiscovery Probe/Discover; UiActivityMonitor, UiBuffering,
  WheelScroller, AppAssets; SetPortalStatusThreadSafe; the RegistrationInput field
  bag; FormConnect.AuthorizeConfiguredChange.

MUST-PRESERVE INVARIANTS
- All public APIs / function signatures other code calls, tests included.
- AGENTS.md non-negotiables: parameterized queries only (the single recorded
  exception is EnsureBtaAppLogin quote-escaping the CREATE LOGIN password into the
  dynamic batch; keep it exactly as it is); one SqlTransaction per multi-entity
  write (registration, routing, user-plus-card-plus-role enrolment, storage moves);
  append-only tbl_AuditTrail with UTC timestamps and JSON snapshots.
- Thread safety: UI controls touched only through Invoke from background threads
  (RFID listener, portal poll, sync timer, UiActivityMonitor tick); replay reads
  and mutations stay under SyncLock(EmbeddedDB.SyncRoot) with SQL I/O outside the
  lock.
- Offline core semantics: XML cache write and read-back with autoincrement
  resync; per-batch replay saves; PendingSync flags; the RowVersion conflict path
  (SYNC_CONFLICT audit plus operator banner); offline DocCode preservation with
  collision fallback; the PullAfterWrite table set.
- UI invariants: HighDpiMode.SystemAware, layout containers instead of fixed
  pixel coordinates, grid format on Shown never in constructors, the
  DataGridStyler.BalanceColumns degenerate-DisplayRectangle guard, combo-box
  selection preserved across sync ticks.
- Auth and session behavior: mirror-based authentication in
  FormMain.AuthenticateUser with SQL-side failed-tap increment, lockout, and
  reset; session timeout via UiActivityMonitor with LOGOUT and SESSION_TIMEOUT
  audit records; the single-instance mutex.
- Settings: DPAPI encryption ("BTA-ENC1:" prefix) loads and re-encrypts
  unchanged; the file format is behavior; the PortalEnabled false default stays.
- Demo-free shipping: demo document seeding stays gated to UseSqlServer = False;
  the self-check stays forced-offline and deterministic on any machine.
- Deliberately simple (leave as-is; changing any of these is behavior work):
  the GetOrCreateByKey read-then-insert race; SessionManager,
  AuthenticationService, and PermissionService unwired while mirror-based auth is
  the live path; the keyboard-wedge typed-UID path; the connected ApplyDirective
  status side effect via DirectiveCodeFor.
- Security-relevant structure may be touched only if the guarantee provably
  survives; when in doubt, report instead of restructure.
- vbproj EmbeddedResource LogicalNames (osgsql.<file>) survive any file move; the
  IDbConnectionFactory seam stays intact.

PRIORITY INSPECTION ORDER
1. God files / god classes: the remaining multi-responsibility candidates are
   FormConnect, EmbeddedDB, DesktopDataCoordinator, and FormDocumentDetail.
   Assess honestly; the big splits already happened, so expect few candidates.
2. Long functions (>40 lines, deep nesting), exempting layout builders as above.
3. Duplicated blocks NOT in the protected list; consolidate into one shared
   function in the same layer.
4. Misleading or generic names; rename toward the domain vocabulary (document,
   routing, directive, storage, enrolment, registry).
5. Dead code, unused parameters, unused settings keys; remove only when
   behavior preservation is provable, otherwise report.
6. Tangled coupling between modules: REPORT only, do not act.

OUTPUT PROTOCOL
- PHASE 1: Propose the full refactor list before touching anything: file,
  technique, why, risk level, ranked by value. WAIT for approval.
- PHASE 2: Execute one approved step at a time. After each: what changed, the
  three gate results, and one line on why behavior is preserved (inspection for
  UI-facing steps, stated as inspection).
- FINAL: Summary: steps completed, net line delta (up or down is fine), what you
  flagged but deliberately left alone and why, and the pre-existing reds that
  remain untouched.

STOP CONDITIONS
Revert and report if: any gate result changes beyond the recorded pre-existing
reds, a step would break a preserved invariant, the suite cannot verify a step,
or a "refactor" would alter observable behavior. Also stop and propose
separately any step that would require editing portal/, db/scripts/, or the
shipped settings JSONs: they are out of scope.
