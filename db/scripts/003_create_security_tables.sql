USE BTA_OSG_DB;
GO

IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tbl_Roles]') AND type in (N'U'))
BEGIN
    CREATE TABLE dbo.tbl_Roles (
        RoleID INT IDENTITY(1,1) PRIMARY KEY,
        RoleCode VARCHAR(50) NOT NULL UNIQUE,
        RoleName NVARCHAR(100) NOT NULL,
        Description NVARCHAR(255) NULL,
        IsSystemRole BIT DEFAULT 0,
        IsActive BIT DEFAULT 1,
        CreatedByUserID INT NULL,
        CreatedAtUTC DATETIME2 DEFAULT SYSUTCDATETIME(),
        ModifiedByUserID INT NULL,
        ModifiedAtUTC DATETIME2 NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tbl_Permissions]') AND type in (N'U'))
BEGIN
    CREATE TABLE dbo.tbl_Permissions (
        PermissionID INT IDENTITY(1,1) PRIMARY KEY,
        PermissionCode VARCHAR(50) NOT NULL UNIQUE,
        PermissionName NVARCHAR(100) NOT NULL,
        Description NVARCHAR(255) NULL,
        IsActive BIT DEFAULT 1
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tbl_RolePermissions]') AND type in (N'U'))
BEGIN
    CREATE TABLE dbo.tbl_RolePermissions (
        RolePermissionID INT IDENTITY(1,1) PRIMARY KEY,
        RoleID INT NOT NULL FOREIGN KEY REFERENCES dbo.tbl_Roles(RoleID),
        PermissionID INT NOT NULL FOREIGN KEY REFERENCES dbo.tbl_Permissions(PermissionID),
        IsActive BIT DEFAULT 1,
        UNIQUE (RoleID, PermissionID)
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tbl_Users]') AND type in (N'U'))
BEGIN
    CREATE TABLE dbo.tbl_Users (
        UserID INT IDENTITY(1,1) PRIMARY KEY,
        Username NVARCHAR(50) NOT NULL UNIQUE,
        FullName NVARCHAR(150) NOT NULL,
        Office NVARCHAR(150) NULL,
        Email NVARCHAR(150) NULL,
        IsActive BIT DEFAULT 1,
        IsLocked BIT DEFAULT 0,
        FailedTapCount INT DEFAULT 0,
        LastFailedTapUTC DATETIME2 NULL,
        CreatedByUserID INT NULL,
        CreatedAtUTC DATETIME2 DEFAULT SYSUTCDATETIME(),
        ModifiedByUserID INT NULL,
        ModifiedAtUTC DATETIME2 NULL,
        RowVersion ROWVERSION
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tbl_UserRoles]') AND type in (N'U'))
BEGIN
    CREATE TABLE dbo.tbl_UserRoles (
        UserRoleID INT IDENTITY(1,1) PRIMARY KEY,
        UserID INT NOT NULL FOREIGN KEY REFERENCES dbo.tbl_Users(UserID),
        RoleID INT NOT NULL FOREIGN KEY REFERENCES dbo.tbl_Roles(RoleID),
        AssignedByUserID INT NULL,
        AssignedAtUTC DATETIME2 DEFAULT SYSUTCDATETIME(),
        IsActive BIT DEFAULT 1,
        UNIQUE (UserID, RoleID)
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tbl_RfidCards]') AND type in (N'U'))
BEGIN
    CREATE TABLE dbo.tbl_RfidCards (
        RfidCardID INT IDENTITY(1,1) PRIMARY KEY,
        UserID INT NOT NULL FOREIGN KEY REFERENCES dbo.tbl_Users(UserID),
        CardPublicID VARCHAR(50) NOT NULL UNIQUE,
        CardLabel NVARCHAR(100) NULL,
        IsActive BIT DEFAULT 1,
        IssuedAtUTC DATETIME2 DEFAULT SYSUTCDATETIME(),
        RevokedAtUTC DATETIME2 NULL,
        RevokedByUserID INT NULL,
        RevocationReason NVARCHAR(255) NULL,
        CreatedByUserID INT NULL,
        CreatedAtUTC DATETIME2 DEFAULT SYSUTCDATETIME(),
        ModifiedByUserID INT NULL,
        ModifiedAtUTC DATETIME2 NULL,
        RowVersion ROWVERSION
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tbl_RfidFailedAttempts]') AND type in (N'U'))
BEGIN
    CREATE TABLE dbo.tbl_RfidFailedAttempts (
        FailedAttemptID BIGINT IDENTITY(1,1) PRIMARY KEY,
        CardPublicID VARCHAR(50) NULL,
        UserID INT NULL FOREIGN KEY REFERENCES dbo.tbl_Users(UserID),
        AttemptAtUTC DATETIME2 DEFAULT SYSUTCDATETIME(),
        MachineName NVARCHAR(100) NULL,
        Reason VARCHAR(100) NOT NULL,
        Success BIT DEFAULT 0
    );
END
GO
