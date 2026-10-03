USE BTA_OSG_DB;
GO

-- Migration 014: Tamper evidence for the audit trail.
-- Each entry stores the SHA-256 of its own columns plus the hash of the entry before it.
-- Sealing is done by the application (AuditRepository), not here: the hash must be
-- computed over values as the server stores them, from one canonical definition.

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'dbo.tbl_AuditTrail') AND name = 'PrevHash')
BEGIN
    ALTER TABLE dbo.tbl_AuditTrail ADD PrevHash VARBINARY(32) NULL;
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'dbo.tbl_AuditTrail') AND name = 'RowHash')
BEGIN
    ALTER TABLE dbo.tbl_AuditTrail ADD RowHash VARBINARY(32) NULL;
END
GO
