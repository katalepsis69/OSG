USE BTA_OSG_DB;
GO

IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tbl_StorageLocations]') AND type in (N'U'))
BEGIN
    CREATE TABLE dbo.tbl_StorageLocations (
        StorageLocationID INT IDENTITY(1,1) PRIMARY KEY,
        CabinetID NVARCHAR(50) NOT NULL,
        ShelfNo NVARCHAR(50) NULL,
        BoxCode NVARCHAR(50) NULL,
        Description NVARCHAR(200) NULL,
        LocationKey AS (CabinetID + '|' + ISNULL(ShelfNo,'') + '|' + ISNULL(BoxCode,'')) PERSISTED,
        IsActive BIT DEFAULT 1,
        CreatedByUserID INT NULL,
        CreatedAtUTC DATETIME2 DEFAULT SYSUTCDATETIME(),
        ModifiedByUserID INT NULL,
        ModifiedAtUTC DATETIME2 NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tbl_Documents]') AND type in (N'U'))
BEGIN
    CREATE TABLE dbo.tbl_Documents (
        DocumentID INT IDENTITY(1,1) PRIMARY KEY,
        DocCode VARCHAR(30) NOT NULL UNIQUE,
        Title NVARCHAR(300) NOT NULL,
        DocumentTypeID INT NULL FOREIGN KEY REFERENCES dbo.tbl_DocumentTypes(DocumentTypeID),
        OriginOffice NVARCHAR(200) NULL,
        DestinationOffice NVARCHAR(200) NULL,
        StatusID INT NULL FOREIGN KEY REFERENCES dbo.tbl_DocumentStatuses(StatusID),
        RegisteredByUserID INT NULL FOREIGN KEY REFERENCES dbo.tbl_Users(UserID),
        RegisteredAtUTC DATETIME2 DEFAULT SYSUTCDATETIME(),
        ReceivedDate DATE NULL,
        CurrentStorageLocationID INT NULL FOREIGN KEY REFERENCES dbo.tbl_StorageLocations(StorageLocationID),
        GoogleDriveUrl NVARCHAR(1000) NULL,
        Remarks NVARCHAR(MAX) NULL,
        IsDeleted BIT DEFAULT 0,
        DeletedByUserID INT NULL,
        DeletedAtUTC DATETIME2 NULL,
        DeletionReason NVARCHAR(255) NULL,
        CreatedByUserID INT NOT NULL,
        CreatedAtUTC DATETIME2 DEFAULT SYSUTCDATETIME(),
        ModifiedByUserID INT NULL,
        ModifiedAtUTC DATETIME2 NULL,
        RowVersion ROWVERSION
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tbl_DocumentSequences]') AND type in (N'U'))
BEGIN
    CREATE TABLE dbo.tbl_DocumentSequences (
        DocumentTypeCode VARCHAR(10) NOT NULL,
        SequenceYear SMALLINT NOT NULL,
        LastNumber INT DEFAULT 0,
        UpdatedAtUTC DATETIME2 DEFAULT SYSUTCDATETIME(),
        PRIMARY KEY (DocumentTypeCode, SequenceYear)
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tbl_DocumentAssignments]') AND type in (N'U'))
BEGIN
    CREATE TABLE dbo.tbl_DocumentAssignments (
        AssignmentID INT IDENTITY(1,1) PRIMARY KEY,
        DocumentID INT NOT NULL FOREIGN KEY REFERENCES dbo.tbl_Documents(DocumentID),
        AssignedUserID INT NULL FOREIGN KEY REFERENCES dbo.tbl_Users(UserID),
        AssignedOffice NVARCHAR(200) NULL,
        AssignedRoleID INT NULL FOREIGN KEY REFERENCES dbo.tbl_Roles(RoleID),
        AssignedByUserID INT NOT NULL FOREIGN KEY REFERENCES dbo.tbl_Users(UserID),
        AssignedAtUTC DATETIME2 DEFAULT SYSUTCDATETIME(),
        IsActive BIT DEFAULT 1,
        Remarks NVARCHAR(500) NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tbl_ActionDirectives]') AND type in (N'U'))
BEGIN
    CREATE TABLE dbo.tbl_ActionDirectives (
        DirectiveID INT IDENTITY(1,1) PRIMARY KEY,
        DocumentID INT NOT NULL FOREIGN KEY REFERENCES dbo.tbl_Documents(DocumentID),
        DirectiveTypeID INT NOT NULL FOREIGN KEY REFERENCES dbo.tbl_DirectiveTypes(DirectiveTypeID),
        DirectiveText NVARCHAR(1000) NULL,
        IssuedByUserID INT NOT NULL FOREIGN KEY REFERENCES dbo.tbl_Users(UserID),
        IssuedAtUTC DATETIME2 DEFAULT SYSUTCDATETIME(),
        IsActive BIT DEFAULT 1,
        SupersededByDirectiveID INT NULL,
        Remarks NVARCHAR(MAX) NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tbl_RoutingLogs]') AND type in (N'U'))
BEGIN
    CREATE TABLE dbo.tbl_RoutingLogs (
        RoutingLogID INT IDENTITY(1,1) PRIMARY KEY,
        DocumentID INT NOT NULL FOREIGN KEY REFERENCES dbo.tbl_Documents(DocumentID),
        FromStatusID INT NULL FOREIGN KEY REFERENCES dbo.tbl_DocumentStatuses(StatusID),
        ToStatusID INT NOT NULL FOREIGN KEY REFERENCES dbo.tbl_DocumentStatuses(StatusID),
        FromOffice NVARCHAR(200) NULL,
        ToOffice NVARCHAR(200) NULL,
        RoutingRemarks NVARCHAR(1000) NULL,
        RoutedByUserID INT NOT NULL FOREIGN KEY REFERENCES dbo.tbl_Users(UserID),
        RoutedAtUTC DATETIME2 DEFAULT SYSUTCDATETIME()
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tbl_DocumentMovements]') AND type in (N'U'))
BEGIN
    CREATE TABLE dbo.tbl_DocumentMovements (
        MovementID INT IDENTITY(1,1) PRIMARY KEY,
        DocumentID INT NOT NULL FOREIGN KEY REFERENCES dbo.tbl_Documents(DocumentID),
        StorageLocationID INT NOT NULL FOREIGN KEY REFERENCES dbo.tbl_StorageLocations(StorageLocationID),
        MovedByUserID INT NOT NULL FOREIGN KEY REFERENCES dbo.tbl_Users(UserID),
        MovedAtUTC DATETIME2 DEFAULT SYSUTCDATETIME(),
        MovementReason NVARCHAR(500) NULL
    );
END
GO
