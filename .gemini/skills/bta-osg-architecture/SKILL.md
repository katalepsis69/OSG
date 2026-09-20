---
name: bta-osg-architecture
description: >
  Master architectural playbook for BTA OSG Document Status Tracking & Monitoring System.
  Use when building, refactoring, debugging, or testing VB.NET, WinForms (.NET 10),
  Microsoft SQL Server ADO.NET repositories, RFID authentication, and MSTest suites.
---

# BTA OSG Document Tracking : Architecture Reference (Path A v2.1)

Core rules live in `AGENTS.md` and `DESIGN.md` (always loaded).
This skill provides the technical implementation reference for the Windows desktop system.

---

## 1. Application Layer (VB.NET & WinForms on .NET 10)

* **Runtime & Target:** .NET 10.0 Windows Desktop (`net10.0-windows`).
* **Language Idioms:** Modern VB.NET with Option Explicit On, Option Strict On.
* **WinForms UI Thread Safety:**
  - Never update UI controls directly from background threads or async tasks.
  - Check `If Me.InvokeRequired Then Me.Invoke(...)` before manipulating controls from async callbacks or hardware listeners.
* **High DPI Scaling:**
  - `Application.SetHighDpiMode(HighDpiMode.SystemAware)` configured at `Program.Main`.
  - Use TableLayoutPanel, FlowLayoutPanel, and Anchor/Dock properties instead of absolute coordinate positioning.
* **Global Error Handling:**
  - `AppGlobalExceptionHandler.Setup()` traps unhandled thread exceptions (`Application.ThreadException`) and domain exceptions (`AppDomain.CurrentDomain.UnhandledException`), logging to `tbl_AuditTrail` before displaying friendly dialogs.

---

## 2. Database & SQL Server ADO.NET Layer

* **Connection Factory:** `IDbConnectionFactory` / `SqlConnectionFactory` reading from `appsettings.json` / `appsettings.Production.json`.
* **Parameterized Queries (Zero SQL Injection):**
  - Never concatenate variables into SQL strings.
  - Always use `SqlCommand.Parameters.AddWithValue` or explicit `SqlParameter` types.
* **Transactions:**
  - Multi-step operations (e.g., routing a document + updating document status + writing routing log + writing audit trail) MUST be bound inside a single `SqlTransaction`.
* **Audit Trail Immutability:**
  - Table `tbl_AuditTrail` is append-only.
  - Every document movement, directive creation, and user login writes an audit record with UTC timestamp, User ID, Action Type, and JSON snapshots.
* **Fallback Mode:**
  - `EmbeddedDB.vb` serves as an in-memory/isolated fallback when SQL Server is unreachable or during unit tests. Production execution must use `Microsoft.Data.SqlClient`.

---

## 3. Hardware & RFID Security

* **Badge Listener:**
  - `RfidInputListener.vb` intercepts card reader keyboard wedge or PC/SC reader inputs.
  - Reader events must sanitize trailing control characters (CR/LF).
* **Lockout Protection:**
  - 5 consecutive failed attempts trigger temporary terminal lockout (`LockoutThreshold = 5`).
* **Card Data Masking:**
  - Audit logs record masked card identifiers (`CardPublicIDMasked`) to prevent badge credential harvesting from logs.

---

## 4. Testing & Pre-Flight Verification

* **Automated Self-Check:**
  ```powershell
  dotnet run --project src/BTA_OSG_DocumentTracking -- /test
  ```
  Verifies DB connection, seed reference data, RFID auth mock, and PDF link validation.
* **MSTest Suite:**
  ```powershell
  dotnet test tests/BTA_OSG_DocumentTracking.Tests
  ```
  Covers AuditService, DirectiveWorkflow, DocumentCoding, PdfLinkValidation, and RBAC policies.

---

## 5. Production Compilation & Publishing

* **Self-Contained Executable:**
  ```powershell
  dotnet publish src/BTA_OSG_DocumentTracking -c Release -r win-x64 --self-contained true
  ```
  Output executable:
  `src/BTA_OSG_DocumentTracking/bin/Release/net10.0-windows/win-x64/publish/BTA_OSG_DocumentTracking.exe`
