USE BTA_OSG_DB;
GO

IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tbl_AuditTrail]') AND type in (N'U'))
BEGIN
    CREATE TABLE dbo.tbl_AuditTrail (
        AuditID BIGINT IDENTITY(1,1) PRIMARY KEY,
        EventAtUTC DATETIME2 DEFAULT SYSUTCDATETIME(),
        SessionID UNIQUEIDENTIFIER NULL,
        UserID INT NULL,
        UsernameSnapshot NVARCHAR(50) NULL,
        FullNameSnapshot NVARCHAR(150) NULL,
        RoleSnapshot NVARCHAR(200) NULL,
        ActionType VARCHAR(100) NOT NULL,
        EntityType VARCHAR(100) NULL,
        EntityID VARCHAR(100) NULL,
        DocumentCode VARCHAR(30) NULL,
        OldValuesJson NVARCHAR(MAX) NULL,
        NewValuesJson NVARCHAR(MAX) NULL,
        MachineName NVARCHAR(100) NULL,
        ClientInfo NVARCHAR(200) NULL,
        Success BIT DEFAULT 1,
        FailureReason NVARCHAR(500) NULL,
        CorrelationId UNIQUEIDENTIFIER NULL,
        CardPublicIDMasked VARCHAR(50) NULL,
        ApplicationVersion NVARCHAR(50) NULL
    );
END
GO
