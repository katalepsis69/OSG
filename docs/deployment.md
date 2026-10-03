# BTA OSG Document Tracking & Monitoring System - Deployment Guide
**Office of the Secretary-General (OSG), Bangsamoro Transition Authority Parliament**
**Version 2.1 (Production Path A)**

---

## 1. System Requirements
- **Target OS**: Windows 10/11 x64 or Windows Server 2022+
- **Runtime**: .NET 10.0 Windows Desktop Runtime
- **Embedded preview**: Microsoft WebView2 Evergreen Runtime (preinstalled on Windows 11 and Windows 10 21H2+; on Windows Server 2022 install it from https://developer.microsoft.com/microsoft-edge/webview2/). Without it the Digital Soft-Copy Preview tab shows a fallback message; "Open in External Browser" keeps working.
- **Database**: Microsoft SQL Server Standard / Enterprise / Express
- **Peripherals**: PC/SC compatible USB RFID Desktop Reader (13.56MHz Mifare / 125kHz EM depending on staff badge specification)

## 2. Database Provisioning

**Zero-scripts flow (primary):** install SQL Server Express on the server PC, then launch the
app there. First run opens the Connect dialog automatically (later via "Reconfigure Station
Setup" in the Admin tab). Type the shared password and press **Connect**: the dialog scans
the network and its own machine for `BTA_OSG_DB`. When the local SQL Server answers but is
missing the schema, the login, or TCP, the **Set up this server** button appears (it only
exists on a machine that hosts SQL Server): it confirms, then switches the instance to Mixed
Mode authentication (a fresh Express install answers Windows-only, which no seat login can
pass), creates `BTA_OSG_DB`, applies schema scripts 001-015 (embedded in the exe, idempotent,
nothing ever dropped), creates the `bta_app` SQL login with db_owner on `BTA_OSG_DB`, enables
TCP on static port 1433, and finishes the OS side under one administrator prompt: restarts
the SQL Server service, starts the SQL Browser service, and opens TCP 1433 and UDP 1434 in
Windows Firewall. When it finishes it re-tests the connection. Every other workstation just
launches the app, types the same shared password, and presses Connect; no script ever runs by
hand and nothing outside the exe ships.

Set up this server must run **on the server machine**: it needs a sysadmin Windows
connection, which in a workgroup exists only on the server console. A seat whose connection
test fails against a remote server is told to confirm the password with whoever set the
server up; it is never offered provisioning.

The `bta_app` password is an office secret: it reads and writes every document. Choose it
once, share it with the seats, and keep `CHECK_EXPIRATION` off (the dialog already does;
an expired login would lock every seat out at once).

**DBA fallback (equivalent result):** run the scripts in `db/scripts/` in numerical order
using SSMS or `sqlcmd`. These files are also the canonical schema source:
1. `001_create_database.sql` - Provisions `BTA_OSG_DB`
2. `002_create_reference_tables.sql` - Document types, statuses, directive types
3. `003_create_security_tables.sql` - Users, roles, permissions, RFID cards
4. `004_create_document_tables.sql` - Documents, sequences, directives, routing, storage
5. `005_create_audit_tables.sql` - Append-only audit trail
6. `006_create_indexes.sql` - Performance and uniqueness indexes
7. `007_create_seed_data.sql` - Initial reference data only (document types, statuses, roles, permissions, directive types); it seeds no user, role assignment, or RFID card
8. `008_osg_target_migration.sql` - Desktop vocabulary alignment (statuses, types, section roles)
9. `009_portal_external_control_number.sql` - Portal external control number column
10. `010_osg_workstation_alignment.sql` - Per-user capability flags (`CanRoute`, `CanMove`, `CanSoftCopy`); the demo staff desk logins, section roles, and cards it once carried were removed
11. `011_create_workstation_heartbeat.sql` - Per-seat heartbeat table behind the Admin tab's Workstation Sync Status grid
12. `012_retire_bootstrap_admin.sql` - Revokes the previously published bootstrap administrator and card on servers that already carry them, and only once another active System Administrator exists, so no server is left with no administrator at all
13. `013_realign_document_sequences.sql` - Raises each `tbl_DocumentSequences` counter to the highest document number already on file, repairing codes minted by an offline seat that never advanced the shared counter
14. `014_audit_hash_chain.sql` - Adds `PrevHash` and `RowHash` to `tbl_AuditTrail`, so every entry seals the entries after it
15. `015_replay_and_integrity_fixes.sql` - Makes `tbl_RoutingLogs.ToStatusID` nullable (non-status offline actions replay with no status), indexes `tbl_DocumentAssignments(DocumentID)` for the snapshot pull, and pins `FlowDirection` to INCOMING/OUTGOING

**Seed policy (final, 2026-10-03).** An office server is never seeded with demo data: the
provisioning chain writes the reference vocabulary only, the staff list starts empty, and
the first administrator is claimed on the station. The SQL integration tests run against a
throwaway `BTA_OSG_DB_TEST` catalog that is dropped when the suite ends, so a test run
cannot leave documents or staff behind in `BTA_OSG_DB`; the offline self-check (`/test`)
seeds at most five fixtures into the local store and removes them before it exits. For a
server that predates this policy, `db/scripts/purge_seed_leftovers.sql` lists and then (after
review) removes leftover probe documents and staff; it is manual and never provisioned.

The audit trail is tamper-evident: changing, removing, or reordering any sealed entry breaks
the chain at that point. `BTA_OSG_DocumentTracking.exe /verify-audit` recomputes the whole chain
and prints the first entry that no longer matches, exiting 0 when every sealed entry is intact.
It reads the database this workstation is configured for, and prints that server name first. If
that server does not answer, it says so and reads a SQL instance on this machine that holds
`BTA_OSG_DB` instead, so the command is useful on a laptop that is not the office server. To name
a server yourself, pass it: `/verify-audit "OSG-SQL-01"`, or a whole connection string when the
saved trust settings do not suit the machine, as with a test server holding a self-signed
certificate. The choice is never written back to the station's settings. Entries older than
script 014 carry no hash until the app seals them, which it does once in the background on the
first connected launch.

The SQL login the app's setup would create (`bta_app`) is its job; a DBA recreating it by
hand should mirror the same settings (`CHECK_POLICY = ON, CHECK_EXPIRATION = OFF`, db_owner
on `BTA_OSG_DB` only). Windows Authentication for the app remains possible in AD offices by
hand-editing the saved connection string (`Integrated Security=true`); the app's setup
configures SQL authentication only, because workgroup seats cannot carry Windows credentials.

## 3. Configuration Setup
Configure `appsettings.Production.json`:
```json
{
  "Database": {
    "ConnectionString": "Server=OSG-SQL-01;Database=BTA_OSG_DB;Integrated Security=true;Encrypt=true;TrustServerCertificate=false;",
    "UseSqlServer": true
  },
  "Rfid": {
    "LockoutThreshold": 5
  },
  "PdfLink": {
    "AllowedHosts": ["drive.google.com", "docs.google.com"],
    "RequireHttps": true
  },
  "Session": {
    "TimeoutMinutes": 15
  }
}
```

Hand-written JSON loads as-is, so this file can be pre-seeded by IT. The load order is
`Resources/appsettings.json`, then `appsettings.{BTA_ENVIRONMENT}.json` (or
`appsettings.Production.json` when no environment variable is set), then the AppData
fallback, then `appsettings.local.json`, so later layers override earlier ones. Once the
Connect dialog saves, it writes `Resources/appsettings.local.json` DPAPI-encrypted
(`BTA-ENC1:` prefix) for the Windows user who ran it, because the connection string carries
the SQL login password. A saved file cannot be copied to another machine or user account;
pre-seed the plain JSON instead, or re-run the setup dialog on the target workstation.

Only a System Administrator badge can open the setup dialog on a configured workstation
(first run on an unconfigured install is open, so a fresh machine can be set up at all).
Changing the station mode there now applies to the running session: the app reloads its
settings, rebuilds its data layer, and re-probes the server, and a single confirmation dialog
states the mode the station is in. No restart is needed for that. A restart is still required
only for the Citizen Web Portal and Data Analytics switches, because the main window reads
those while it builds its navigation and poll timers, and the confirmation says so when one of
those changed.

When Connect runs, it enumerates the SQL Servers the Browser service announces plus the
instances installed on the local machine (read from the registry, so a fresh Express install
is found even before TCP and the Browser service exist), probes each for `BTA_OSG_DB` with
the three sentinel tables using the shared SQL login, and connects to the first complete
match. A server that answers but rejects the login is reported as such: on the server
machine the dialog offers Set up this server, on a seat it says to confirm the password with
whoever set the server up. The scan never saves anything by itself; saving happens only
through a successful Connect, and when the network is silent (Browser service off, firewall
blocking UDP 1434) the server address stays manual in the options panel.

## 3.1 Multi-Workstation Synchronization

Every workstation points at the same `BTA_OSG_DB`. When the connection string is live
(`UseSqlServer: true`) SQL Server is the single source of truth:

- Writes go to SQL Server, and the local XML cache is refreshed from SQL Server.
- Workstations poll SQL Server on the configured interval (`Database.SyncIntervalSeconds`,
  default 10, floored at 5) plus a small one-time jitter, so a document registered at the
  Records desk appears on the Secretary-General's screen without any manual refresh and a
  room full of machines booted at 08:00 does not poll in one wave.
- Each seat upserts a heartbeat row per pull (`tbl_WorkstationHeartbeat`: app version,
  pending offline records, last error). The Admin tab's **Workstation Sync Status** grid
  shows every seat; one that has been silent past three intervals reads Stale, which is how
  an offline machine is meant to look.
- An offline edit replays with an optimistic guard: if another seat moved the document in
  the meantime (rowversion mismatch), the server's newer copy stands, the local edit is
  discarded, a `SYNC_CONFLICT` audit entry records it, and the status banner says so. All
  other replayed work (directives, routing, movements, new registrations) is append-only and
  unaffected.
- The local XML file (`bta_osg_db.xml`) is a cache, not an authority, while SQL Server is
  reachable. Editing or deleting it does not change shared data, and the next sync overwrites it.
- Two users registering a document at the same second get two distinct document codes:
  codes are minted from `tbl_DocumentSequences` inside a locking transaction.

When the database is unreachable (`UseSqlServer: false`, or a host that does not answer the
startup probe) the desktop behaves exactly as it does today: the XML cache is the authority,
nothing polls, and no dialog interrupts the desk. Actions taken in that state stay local until
the database returns, at which point the next sync mirrors SQL Server onto the cache.

**What a fresh workstation contains:** an empty document registry, an empty staff list, and
no known card numbers, in either profile. The seed carries the reference vocabulary only
(types, statuses, roles, permissions, directive types, storage locations), never a user, a
role assignment, or an RFID card, so nobody can sign in until an administrator is claimed. On
a station that has no System Administrator the app opens the **Claim the first administrator**
dialog before the main window: the operator names the administering officer, taps that
officer's badge or types its card ID, and confirms the desk or department. In office mode the
claim writes to SQL Server; in the standalone demo profile (`UseSqlServer: false`) it writes
to the local embedded cache. Closing that dialog without claiming ends the application,
because the station would be left with no way to sign in, and a card already enrolled for a
different person is refused, so a claim cannot silently rename and promote an existing staff
member. Enrol the remaining staff badges from the Admin tab. The login window lists the
identities enrolled on that station, so a seat with no reader yet can still sign in; the audit
entry marks such a sign in as an on-screen selection rather than a badge tap, and the list is
empty until a claim or enrolment puts an identity in it. A reader works with no configuration:
it is a keyboard wedge, and a tap submits the card read rather than the highlighted entry. A
server provisioned by an older
release may still carry the bootstrap administrator that release seeded; script
`012_retire_bootstrap_admin.sql` revokes that user and card, but only once some other active
System Administrator exists, because removing the last administrator would lock every
workstation out. The shipped `Resources/appsettings.json` / `appsettings.Production.json`
carry no secrets: the portal bridge key is empty and is generated by the setup dialog the
first time the portal is enabled.

## 4. Compilation & Publishing
The ship artifact is what `publish.bat` produces: run it and ship `dist\BTA_OSG_DocumentTracking.exe`
plus `dist\Resources\` (settings live in the vbproj Release configuration, so the batch
carries the whole story: win-x64, self-contained, single file, loose Resources).

Underlying command, for reference (the batch adds the single-file and Resources handling
that run.bat and the setup guide assume):
```powershell
dotnet publish src/BTA_OSG_DocumentTracking -c Release -r win-x64 --self-contained true
```
A raw publish without the batch lands in
`src/BTA_OSG_DocumentTracking/bin/Release/net10.0-windows/win-x64/publish/`, which
`run.bat` will not launch; ship `dist\` instead.

## 5. Pre-Flight Verification
Run automated self-check suite:
```powershell
dotnet run --project src/BTA_OSG_DocumentTracking -- /test
```
Exit code 0 confirms all 14 validation milestones passed.

## 6. Backups
**SQL Server (the authoritative database):** `db/scripts/backup_sqlserver.sql` dumps
`BTA_OSG_DB` to the server's default backup directory with CHECKSUM and verifies the
result with RESTORE VERIFYONLY. Schedule it daily: either a SQL Server Agent job, or a
Windows Task Scheduler task running
`sqlcmd -S <host> -E -b -i "C:\path\to\db\scripts\backup_sqlserver.sql"` (the `-b` flag
turns the script's failure RAISERROR into a nonzero exit code).
A failed backup must be noticed: check the task's exit code and the SQL Server log. The
script writes into the instance's default backup directory under a `BTA_OSG\` subfolder
and prunes that subfolder's own files after 30 days. A daily full backup leaves a
worst-case 24-hour data-loss window; tighten it with hourly log backups under the FULL
recovery model if the office needs less.

**Offline work on each workstation** lives in the XML cache next to the exe and replays
to SQL Server on reconnect; it also survives application restarts (the cache is reloaded
at startup). The cache is still only a mirror while SQL Server is reachable.

**Portal (when enabled):** `portal/deploy/backup-portal.sh` dumps the portal MySQL
database daily; `portal/cron/crontab.txt` includes the schedule line. On a Windows
desk the portable PHP and MySQL under `portal\runtime` are started by
`PortalServerManager`, along with the optional `portal\cloudflared.exe` tunnel.
[`docs/portal-deployment-runbook.md`](./portal-deployment-runbook.md) covers the
alternative of hosting the portal on a cloud VPS.
