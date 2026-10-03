USE BTA_OSG_DB;
GO

-- Desktop permission flags. The desktop Admin screen owns these three checkboxes, and the
-- Users mirror reads them so an unchecked box is not silently reset by the next sync.
IF COL_LENGTH('dbo.tbl_Users', 'CanRoute') IS NULL
    ALTER TABLE dbo.tbl_Users ADD CanRoute BIT NOT NULL CONSTRAINT DF_tbl_Users_CanRoute DEFAULT 1;
IF COL_LENGTH('dbo.tbl_Users', 'CanMove') IS NULL
    ALTER TABLE dbo.tbl_Users ADD CanMove BIT NOT NULL CONSTRAINT DF_tbl_Users_CanMove DEFAULT 1;
IF COL_LENGTH('dbo.tbl_Users', 'CanSoftCopy') IS NULL
    ALTER TABLE dbo.tbl_Users ADD CanSoftCopy BIT NOT NULL CONSTRAINT DF_tbl_Users_CanSoftCopy DEFAULT 1;
GO

-- Desk logins, section roles, cards, and office assignments are never seeded: a provisioned
-- server holds no identity at all until the first run claims the System Administrator, and
-- every other desk is enrolled from the Admin tab.
