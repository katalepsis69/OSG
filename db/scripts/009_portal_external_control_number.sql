USE BTA_OSG_DB;
GO

-- Migration 009: Add ExternalControlNumber to tbl_Documents
-- Links internal document records to external public portal tracking numbers.
-- Filtered unique index prevents duplicate ingestion of external submissions.

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'dbo.tbl_Documents') AND name = 'ExternalControlNumber')
BEGIN
    ALTER TABLE dbo.tbl_Documents ADD ExternalControlNumber VARCHAR(24) NULL;
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.tbl_Documents') AND name = 'UQ_tbl_Documents_ExternalControlNumber')
BEGIN
    CREATE UNIQUE NONCLUSTERED INDEX UQ_tbl_Documents_ExternalControlNumber
    ON dbo.tbl_Documents(ExternalControlNumber)
    WHERE ExternalControlNumber IS NOT NULL;
END
GO
