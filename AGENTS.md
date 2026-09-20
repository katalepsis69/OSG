# AGENTS.md : BTA OSG Document Tracking (Path A v2.1)

## 1. Non-Negotiable Core Rules

1. **VB.NET & Modern .NET 10 Desktop:**
   - Target Framework: `net10.0-windows` with Windows Forms (`<UseWindowsForms>true</UseWindowsForms>`).
   - SDK-style project file with modern `.slnx` solution format.
   - Code must maintain `Option Explicit On` and `Option Strict On` hygiene.

2. **WinForms UI Thread Safety & High DPI:**
   - Never touch UI controls directly from background threads or async hardware listeners.
   - Always check `If Me.InvokeRequired Then Me.Invoke(...)` before manipulating controls.
   - Support `HighDpiMode.SystemAware` scaling. Avoid hardcoded pixel coordinates; use layout containers (`TableLayoutPanel`, `FlowLayoutPanel`, docking/anchoring).

3. **Microsoft SQL Server ADO.NET Architecture:**
   - Target Database: `BTA_OSG_DB` (SQL Server Standard/Enterprise/Express).
   - Driver: `Microsoft.Data.SqlClient` (v6.0.1+).
   - **Parameterized Queries Mandatory:** Never concatenate variables into SQL strings. Always use `SqlParameter` to prevent SQL injection.
   - **Atomic Transactions:** Multi-entity operations (routing updates, status transitions, storage movements) must be wrapped in a single `SqlTransaction`.
   - **Offline / Test Fallback:** `EmbeddedDB.vb` is for isolated smoke testing and demo fallback only; production paths must use `SqlConnectionFactory`.

4. **Append-Only Audit Trail:**
   - Table `tbl_AuditTrail` is append-only.
   - Every document registration, directive issuance, routing step, and user authentication MUST write an audit record with UTC timestamp, User ID, Action Type, and JSON snapshots.

5. **RFID & Hardware Security:**
   - RFID reader listener intercepts card swipes with trailing control character sanitization.
   - Lockout threshold: 5 consecutive failed card reads trigger terminal lockout.
   - Masked card public ID logged in audit trail; never log raw card credential tokens.

6. **Testing & Pre-Flight Gates:**
   - Self-Check Suite: `dotnet run --project src/BTA_OSG_DocumentTracking -- /test` must exit with code 0.
   - MSTest Suite: `dotnet test tests/BTA_OSG_DocumentTracking.Tests` must be 100% passing before merges.

---

## 2. Installed Agent Skills Inventory (18 Skills)

Skills are installed in [`.agents/skills/`](./.agents/skills/) and [`.gemini/skills/`](./.gemini/skills/):

- **Core Architecture:** `bta-osg-architecture` (Path A v2.1 technical playbook)
- **.NET & VB.NET Performance:** `analyzing-dotnet-performance`, `dotnet-backend-patterns`, `dotnet-pinvoke`, `msbuild-antipatterns`, `copy-to-output-directory`, `directory-build-organization`, `check-bin-obj-clash`
- **Database & SQL Server:** `azure-sql-best-practices`, `sql-optimizer`
- **Security & Threat Defense:** `security-threat-model`, `security-best-practices`
- **Testing & Quality Assurance:** `exp-test-maintainability`, `code-review`, `thermo-nuclear-code-quality-review`, `gh-fix-ci`
- **Desktop Usability & Clean UI:** `antislop`, `antislop-human`

---

## 3. Project References

- **Design System:** [`DESIGN.md`](./DESIGN.md) (Desktop WinForms UI/UX rules)
- **Deployment Guide:** [`docs/deployment.md`](./docs/deployment.md)
- **Production Plan:** [`doc_text.txt`](./doc_text.txt)
- **Database Scripts:** [`db/scripts/`](./db/scripts/) (001 to 007)
