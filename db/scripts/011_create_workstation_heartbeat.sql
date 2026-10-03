-- BTA OSG Document Tracking: workstation heartbeat. Applied by the app's embedded
-- provisioning (scripts 001-015 run verbatim via Set up this server), or by hand here
-- for DBA-managed servers.
-- Each seat upserts one row per sync pull; the desktop Admin tab's Workstation Sync Status
-- grid reads it so the administrator can see which seats are alive and which fell offline.

USE BTA_OSG_DB;
GO

IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tbl_WorkstationHeartbeat]') AND type in (N'U'))
BEGIN
    CREATE TABLE dbo.tbl_WorkstationHeartbeat (
        MachineName SYSNAME NOT NULL CONSTRAINT PK_WorkstationHeartbeat PRIMARY KEY,
        LastSyncUTC DATETIME2 NOT NULL CONSTRAINT DF_WorkstationHeartbeat_LastSyncUTC DEFAULT SYSUTCDATETIME(),
        AppVersion NVARCHAR(50) NULL,
        PendingOutbox INT NOT NULL CONSTRAINT DF_WorkstationHeartbeat_PendingOutbox DEFAULT 0,
        LastError NVARCHAR(500) NULL
    );
END
GO
