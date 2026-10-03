# BTA OSG Document Tracking - Setup Guide

Quick reference for setting up on a new PC or office network. Covers two scenarios:
1. **Thesis Demo Setup** - Desktop app + web portal on one PC, using the portal runtime bundled in `portal\runtime\` (no Docker).
2. **Office LAN Setup** - Multi-user desktop app over local network with SQL Server Express (portal disabled).

---

## Part 1: Thesis Demo Setup (Desktop + Web Portal)

Runs both the desktop app and the public web portal on a single Windows PC for defense or evaluation.

### Prerequisites
- Windows 10 or 11 (64-bit)
- **Zero External Installs Needed**: Portable PHP 8.3 and MySQL 8.4 runtime, with the portal database already built and holding no submissions, sit inside `portal\runtime\`. (External Laragon / XAMPP on `C:\` are completely optional fallbacks).

---

### Step 1: Copy Project Folders to the PC
Copy these folders to the test machine (or keep on a flash drive):
- `dist\` (the compiled desktop application)
- `portal\` (the self-contained web portal with bundled portable runtime)

---

### Step 2: Launch the Portal & Database (Zero Configuration)

You do **not** need to install or configure anything on `C:\`. The portal is completely self-contained.

#### Method 1: Automatic 1-Click from the Desktop App (Recommended)
1. Run `dist\BTA_OSG_DocumentTracking.exe`.
2. Go to the **Portal Intake** tab.
3. Click the **`Open OSGPortal`** button.
   - The desktop app automatically starts the bundled MySQL server and PHP web portal in the background.
   - Your default browser opens immediately to `http://localhost:8085`.
   - When the desktop app closes, the background services stop cleanly.

#### Method 2: Standalone Launcher Script
1. Open the `portal\` folder.
2. Double-click `start-portal.bat`.
   - It boots the bundled MySQL server and PHP server automatically.
   - Keep the window open during testing.

#### Method 3: Mobile & Outside Access (Cloudflare Tunnel)
To access the portal from smartphones, tablets, or external networks without port forwarding:
1. **From Desktop App:** Open the **Portal Intake** tab and click **`☁️ Cloudflare Tunnel`**.
2. **Or Standalone:** Double-click `portal\start-tunnel.bat`.
3. A console window displays an instant public HTTPS URL (e.g. `https://<unique-id>.trycloudflare.com`).
4. Open the link on any phone or device. Visitors connect directly with zero reminder pages, passwords, or router setup.

---

### Step 3: Configure Desktop App for Demo Mode
Open `dist\Resources\appsettings.json` in Notepad and verify the following values:

```json
{
  "Database": {
    "UseSqlServer": false
  },
  "Portal": {
    "PortalEnabled": true,
    "BaseUrl": "http://localhost:8085",
    "BridgeKey": ""
  }
}
```

- `"UseSqlServer": false` tells the app to use the bundled standalone XML database (`dist\bta_osg_db.xml`). **No SQL Server install needed.**
- Leave `Portal.BridgeKey` empty in the shipped file: the app generates a fresh key the first time the portal is enabled from Station Setup (the historically published demo key is treated as burned). The same key belongs in `portal\config\secrets.local` as `BRIDGE_SECRET_KEY=...` so the manual launcher matches; that file is gitignored.
- The login window lists the identities this station has actually enrolled under **Sign in as:**, and signs in the one you pick. There is nothing to pick until an administrator is claimed, and no fixed or published badge number exists anywhere in the product.
- A physical reader still works with no configuration: it is a keyboard wedge, so the window collects its keystrokes and the reader's trailing Enter submits the tapped card in preference to whatever is highlighted. A sign in picked from the list is audited as an on-screen selection rather than a badge tap.

---

### Step 4: Run and Verify
1. Double-click `dist\BTA_OSG_DocumentTracking.exe`.
2. The demo store ships with no accounts, so the app opens the **Claim the first administrator** dialog before the main window. Complete it as described under "First Run: Claim the Administrator" below, then sign in by tapping that same card.
3. Open `http://localhost:8085` in your browser: citizen submission and tracking pages are live.
4. In the desktop app, go to **Portal Intake** to review and import pending submissions.

*Note on OTP emails:* The portal sends the code through the Brevo API, whose key comes from `portal\config\secrets.local` (manual launcher) or the desktop's encrypted settings (app-launched portal). A submission only reaches **Portal Intake** after its code is entered, because the queue is built from verified documents and a document is created at that moment. So if the code email is not arriving, test intake with a mailbox that can actually receive it. No phpMyAdmin ships with the bundled runtime.

#### First Run: Claim the Administrator
Nothing is seeded as a person any more. The app starts with an empty document registry, an empty staff list, and no known card numbers, so nobody can sign in until an administrator is claimed on that station.

1. On the **Claim the first administrator** dialog, enter the officer's **Full name**.
2. Fill **RFID card UID** by tapping the badge on the USB reader, or typing the card ID and pressing Enter.
3. Fill **Desk or department**; it defaults to `ICT / Systems Administration`.
4. Click **Claim administrator**. In demo mode the claim is written to the local embedded cache; in office mode it is written to SQL Server.
5. Tap the same card at the login window to reach the main window.

Two things to know about this step:
- Closing the dialog without claiming ends the application, because the station would be left with no way to sign in.
- The claim only enrols a spare card. A card that is already enrolled for a different person is refused, so claiming cannot rename and promote an existing staff member.

#### Enrolling Staff (no pre-configured accounts)
There is no list of ready-made accounts or card numbers to work from. After signing in as the claimed administrator:

1. Open the **User & RFID Admin** tab.
2. Register each officer: full name, system role, desk or section, and the RFID card UID of the badge issued to them.
3. Record which physical badge was issued to which officer as you enrol it, and keep that record with the station setup notes. The registry only ever holds the card IDs you enter there.

---
---

## Part 2: Office LAN Setup (Production / Post-Thesis)

Permanent office setup. SQL Server runs on one host PC; staff members run the desktop application from their workstations over LAN. The web portal is completely disabled.

```text
[ Server / Host PC ]
  SQL Server 2022 Express (BTA_OSG_DB)
  IP: e.g., 192.168.1.50
        |
        +--- Workstation 1 (dist\BTA_OSG_DocumentTracking.exe)
        +--- Workstation 2 (dist\BTA_OSG_DocumentTracking.exe)
        +--- Workstation 3 (dist\BTA_OSG_DocumentTracking.exe)
```

No web server, PHP, XAMPP, or Docker needed.

---

### Step 1: Install SQL Server on the Host PC
On the designated server or host PC:
1. Download and run the free **[SQL Server 2022 Express Installer](https://www.microsoft.com/en-us/sql-server/sql-server-downloads)** (choose Basic).
2. Download and install **[SQL Server Management Studio (SSMS)](https://learn.microsoft.com/sql/ssms/download-sql-server-management-studio-ssms)**.

---

### Step 2: Enable TCP/IP & Windows Firewall
By default, SQL Server Express rejects remote network connections.

#### Enable TCP/IP:
1. Open **SQL Server 2022 Configuration Manager**.
2. Expand **SQL Server Network Configuration** -> **Protocols for SQLEXPRESS**.
3. Right-click **TCP/IP** and select **Enable**.
4. Double-click **TCP/IP**, open the **IP Addresses** tab:
   - Scroll to the bottom section labeled **IPAll**.
   - Clear **TCP Dynamic Ports** (leave it completely empty).
   - Set **TCP Port** to `1433`.
   - Click **OK**.
5. In the left pane, click **SQL Server Services**, right-click **SQL Server (SQLEXPRESS)**, and click **Restart**.

#### Allow Firewall Inbound Port:
Open PowerShell as Administrator on the host PC and run:
```powershell
New-NetFirewallRule -DisplayName "SQL Server (TCP 1433)" -Direction Inbound -Protocol TCP -LocalPort 1433 -Action Allow
```

---

### Step 3: Run Database Scripts
Open SSMS, connect to `localhost\SQLEXPRESS`, and execute the following SQL scripts from the `db\scripts\` folder in this exact sequence:

1. `001_create_database.sql`
2. `002_create_reference_tables.sql`
3. `003_create_security_tables.sql`
4. `004_create_document_tables.sql`
5. `005_create_audit_tables.sql`
6. `006_create_indexes.sql`
7. `007_create_seed_data.sql`
8. `008_osg_target_migration.sql`
9. `009_portal_external_control_number.sql`
10. `010_osg_workstation_alignment.sql`
11. `011_create_workstation_heartbeat.sql`
12. `012_retire_bootstrap_admin.sql`
13. `013_realign_document_sequences.sql`
14. `014_audit_hash_chain.sql`

*(In SSMS: Open file -> press **F5** to run).*

---

### Step 4: Enable SQL Auth & Create App User
In SSMS:
1. Right-click the server name at the top of Object Explorer -> **Properties** -> **Security**.
2. Select **SQL Server and Windows Authentication mode** -> click **OK**.
3. Restart the SQL Server instance (right-click server name -> **Restart**).
4. Open a **New Query** window and execute:

```sql
USE [master];
GO
CREATE LOGIN [bta_app] WITH PASSWORD = 'StrongOfficePassword123!', CHECK_POLICY = ON;
GO
USE [BTA_OSG_DB];
GO
CREATE USER [bta_app] FOR LOGIN [bta_app];
ALTER ROLE [db_datareader] ADD MEMBER [bta_app];
ALTER ROLE [db_datawriter] ADD MEMBER [bta_app];
GO
```

---

### Step 5: Deploy to Client Workstations
Find the host PC's local LAN IP address (open CMD on the host PC and type `ipconfig`, e.g., `192.168.1.50`).

On each staff workstation:
1. Copy the `dist\` folder to the PC (e.g., `C:\BTA_OSG\`).
2. Open `dist\Resources\appsettings.json` (or `appsettings.Production.json`) in Notepad and configure:

```json
{
  "Database": {
    "ConnectionString": "Server=192.168.1.50,1433;Database=BTA_OSG_DB;User Id=bta_app;Password=StrongOfficePassword123!;Encrypt=true;TrustServerCertificate=true;Max Pool Size=100;Connect Timeout=15;",
    "UseSqlServer": true
  },
  "Rfid": {
    "LockoutThreshold": 5
  },
  "Session": {
    "TimeoutMinutes": 15
  },
  "Portal": {
    "PortalEnabled": false
  }
}
```

3. Plug the USB RFID reader into the workstation.
4. Launch `BTA_OSG_DocumentTracking.exe` on the workstation that will administer the office. Provisioning seeds reference data only (document types, statuses, roles, permissions, directive types), so the staff list is empty and the **Claim the first administrator** dialog opens before the main window. Name the administering officer, tap or type their card ID, confirm the desk, and click **Claim administrator**; in office mode the claim is written to SQL Server.
5. Sign in with that card, then enrol each staff member from the **User & RFID Admin** tab and record the badge issued to each.
6. Staff members tap their physical RFID card badge to log in.

---

### Step 6: Multi-LAN Sync & Offline Fallback Mechanics

The desktop application coordinates state across all office LAN workstations:
- **10-Second Background Polling:** Every 10 seconds, client workstations fetch updates from SQL Server. New documents registered at Records Section or directives issued by the Secretary-General appear on all section desks without manual page refreshes.
- **Concurrent Sequence Minting:** Document numbers (`DocCode`) use transaction locks (`UPDLOCK, HOLDLOCK`) on `tbl_DocumentSequences`. Multiple desks registering simultaneously will never generate duplicate codes.
- **Seamless Offline Fallback:** If LAN connectivity drops or the SQL host restarts, the app automatically switches to local cache mode (`EmbeddedDB`). Staff continue registering, routing, and shelving documents without interruption or error dialogs.
- **Automatic Outbox Replay:** Once LAN connection is restored, pending offline records (`PendingSync = True`) replay to SQL Server. New offline documents receive official SQL identities, dependent child records (directives, routing steps, storage moves, audit records) are remapped to the new IDs, and the local cache updates with central state.

---

### Step 7: Workstation Pre-Flight Verification

Verify database connectivity and hardware readiness on any workstation:
```powershell
cd C:\BTA_OSG
.\BTA_OSG_DocumentTracking.exe /test
```
A correctly connected LAN workstation reports:
```text
[PASS] 1. Database connection verified (SQL Server).
...
[PASS] 12. SQL mirror read handled (live database: True).
=========================================================
ALL 14 SELF-CHECK TESTS PASSED SUCCESSFULLY!
=========================================================
```

---

## Configuration Reference Summary

| Setting Key | Thesis Demo | Office LAN (Production) | Description |
|---|---|---|---|
| `Database.UseSqlServer` | `false` | `true` | `false` uses local XML (`EmbeddedDB`); `true` uses SQL Server connection string as primary authority. |
| `Database.ConnectionString` | N/A | `Server=...;Database=BTA_OSG_DB;...` | ADO.NET connection string pointing to the LAN SQL Server host. |
| `Portal.PortalEnabled` | `true` | `false` | Enables/disables portal background polling and queue intake. |
| `Rfid.LockoutThreshold` | `5` | `5` | Failed badge swipe limit before the 5-minute terminal lockout. |
| `Session.TimeoutMinutes` | `15` | `15` | Idle session timeout in minutes. |

---

## Troubleshooting Guide

### 1. Workstation Cannot Connect to SQL Server
- **Check TCP/IP:** Ensure TCP/IP is enabled on port 1433 in **SQL Server Configuration Manager** -> **SQL Server Network Configuration** -> **Protocols for SQLEXPRESS**.
- **Check Firewall:** Test TCP port 1433 reachability from the client workstation:
  ```powershell
  Test-NetConnection -ComputerName 192.168.1.50 -Port 1433
  ```
  If `TcpTestSucceeded : False`, verify the Windows Firewall rule on the host machine.
- **Check SQL Authentication:** If encountering `Login failed for user 'bta_app'`, verify SQL Server is set to Mixed Mode Authentication (SSMS -> Server Properties -> Security -> SQL Server and Windows Authentication mode) and restart SQL Server.

### 2. Portal Bridge Authentication Error (HTTP 401)
- The desktop holds its generated key in `Portal.BridgeKey` (appsettings, encrypted on save); the manual portal launcher reads the same value from `portal\config\secrets.local` as `BRIDGE_SECRET_KEY=...`. Both sides must carry the identical key: if one was regenerated without the other, every import or status push answers 401. There is no published default key.

### 3. USB RFID Reader Not Detected
- Ensure the reader operates as a standard PC/SC or USB HID keyboard emulation device (13.56 MHz Mifare / 125 kHz EM).
- The reader is a keyboard wedge, so a tap registers wherever the login window has focus: its keystrokes are collected by the window and its trailing Enter submits the tapped card.
- With no reader on a station, open the **Sign in as:** list and pick the officer. The list is empty until the administrator claim is saved, and it only ever holds identities enrolled on that station.


