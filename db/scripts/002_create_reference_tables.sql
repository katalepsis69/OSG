USE BTA_OSG_DB;
GO

IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tbl_DocumentTypes]') AND type in (N'U'))
BEGIN
    CREATE TABLE dbo.tbl_DocumentTypes (
        DocumentTypeID INT IDENTITY(1,1) PRIMARY KEY,
        TypeCode VARCHAR(10) NOT NULL UNIQUE,
        TypeName NVARCHAR(100) NOT NULL,
        Prefix VARCHAR(10) NOT NULL,
        IsActive BIT DEFAULT 1,
        SortOrder INT DEFAULT 0
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tbl_DocumentStatuses]') AND type in (N'U'))
BEGIN
    CREATE TABLE dbo.tbl_DocumentStatuses (
        StatusID INT IDENTITY(1,1) PRIMARY KEY,
        StatusCode VARCHAR(50) NOT NULL UNIQUE,
        StatusName NVARCHAR(100) NOT NULL,
        SortOrder INT DEFAULT 0,
        IsActive BIT DEFAULT 1
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tbl_DirectiveTypes]') AND type in (N'U'))
BEGIN
    CREATE TABLE dbo.tbl_DirectiveTypes (
        DirectiveTypeID INT IDENTITY(1,1) PRIMARY KEY,
        DirectiveCode VARCHAR(50) NOT NULL UNIQUE,
        DirectiveName NVARCHAR(150) NOT NULL,
        ResultStatusID INT NULL FOREIGN KEY REFERENCES dbo.tbl_DocumentStatuses(StatusID),
        IsActive BIT DEFAULT 1
    );
END
GO
