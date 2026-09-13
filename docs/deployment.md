# BTA OSG Document Tracking & Monitoring System - Deployment Guide
**Office of the Secretary-General (OSG), Bangsamoro Transition Authority Parliament**
**Version 2.1 (Production Path A)**

---

## 1. System Requirements
- **Target OS**: Windows 10/11 x64 or Windows Server 2022+
- **Runtime**: .NET 10.0 Windows Desktop Runtime (or .NET 8.0 Windows Desktop Runtime)
- **Database**: Microsoft SQL Server Standard / Enterprise / Express
- **Peripherals**: PC/SC compatible USB RFID Desktop Reader (13.56MHz Mifare / 125kHz EM depending on staff badge specification)

## 2. Database Provisioning
Run database scripts in `db/scripts/` in numerical order using SQL Server Management Studio (SSMS) or `sqlcmd`:
1. `001_create_database.sql` - Provisions `BTA_OSG_DB`
2. `002_create_reference_tables.sql` - Document types, statuses, directive types
3. `003_create_security_tables.sql` - Users, roles, permissions, RFID cards
4. `004_create_document_tables.sql` - Documents, sequences, directives, routing, storage
5. `005_create_audit_tables.sql` - Append-only audit trail
6. `006_create_indexes.sql` - Performance and uniqueness indexes
7. `007_create_seed_data.sql` - Initial reference and staff credential seed

## 3. Configuration Setup
Configure `appsettings.Production.json`:
```json
{
  "Database": {
    "ConnectionString": "Server=OSG-SQL-01;Database=BTA_OSG_DB;Integrated Security=true;Encrypt=true;TrustServerCertificate=false;",
    "UseSqlServer": true
  },
  "Rfid": {
    "SimulatorEnabled": false,
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

## 4. Compilation & Publishing
Self-contained publish command (embeds .NET runtime):
```powershell
dotnet publish src/BTA_OSG_DocumentTracking -c Release -r win-x64 --self-contained true
```

Output executable directory:
```text
src/BTA_OSG_DocumentTracking/bin/Release/net10.0-windows/win-x64/publish/BTA_OSG_DocumentTracking.exe
```

## 5. Pre-Flight Verification
Run automated self-check suite:
```powershell
dotnet run --project src/BTA_OSG_DocumentTracking -- /test
```
Exit code 0 confirms all 12 validation milestones passed.
