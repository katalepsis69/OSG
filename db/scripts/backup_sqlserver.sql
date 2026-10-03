-- BTA OSG Document Tracking & Monitoring System: SQL Server Backup Script
-- Creates a full compressed database backup with CHECKSUM and VERIFYONLY, prunes
-- backups older than 30 days, and fails loudly: the final RAISERROR uses state 127,
-- which makes sqlcmd exit nonzero, so a scheduled task reports the failure.
-- Schedule via Task Scheduler: sqlcmd -S <host> -E -b -i backup_sqlserver.sql
-- The backup lands in the instance's default backup directory under BTA_OSG\, so
-- pruning never touches other databases' files.

USE master;
GO

DECLARE @DatabaseName NVARCHAR(128) = N'BTA_OSG_DB';
DECLARE @SubDirectory NVARCHAR(260);
DECLARE @BackupDirectory NVARCHAR(260) = CONVERT(NVARCHAR(260), SERVERPROPERTY('InstanceDefaultBackupPath'));

IF @BackupDirectory IS NULL OR @BackupDirectory = N''
BEGIN
    DECLARE @RegDirectory NVARCHAR(260);
    EXEC master.dbo.xp_instance_regread
        N'HKEY_LOCAL_MACHINE', N'SOFTWARE\Microsoft\MSSQLServer\MSSQLServer',
        N'BackupDirectory', @RegDirectory OUTPUT;
    SET @BackupDirectory = ISNULL(@RegDirectory, N'');
END

IF @BackupDirectory IS NULL OR @BackupDirectory = N''
    RAISERROR('Backup failed: could not resolve the instance backup directory. Set @BackupDirectory manually in this script.', 16, 127) WITH LOG;
ELSE
BEGIN
    SET @SubDirectory = @BackupDirectory
        + CASE WHEN RIGHT(@BackupDirectory, 1) = N'\' THEN N'' ELSE N'\' END
        + N'BTA_OSG\';

    BEGIN TRY
        EXEC master.dbo.xp_create_subdir @SubDirectory;
    END TRY
    BEGIN CATCH
        RAISERROR('Backup failed: cannot create the BTA_OSG backup subdirectory.', 16, 127) WITH LOG;
        RETURN;
    END CATCH;

    BEGIN TRY
        DECLARE @Cutoff DATETIME = DATEADD(DAY, -30, GETDATE());
        EXEC master.dbo.xp_delete_file 0, @SubDirectory, N'bak', @Cutoff, 0;
        PRINT 'Pruned BTA_OSG backups older than 30 days (if any).';
    END TRY
    BEGIN CATCH
        PRINT 'Retention pruning could not run (non-fatal): ' + ERROR_MESSAGE();
    END CATCH;

    DECLARE @Timestamp VARCHAR(20) = CONVERT(VARCHAR(20), GETDATE(), 112) + '_' + REPLACE(CONVERT(VARCHAR(8), GETDATE(), 108), ':', '');
    DECLARE @BackupPath NVARCHAR(560) = @SubDirectory + @DatabaseName + N'_FULL_' + @Timestamp + N'.bak';
    -- Express (EngineEdition 4) rejects WITH COMPRESSION outright (error 1844), and the
    -- office server runs Express, so compression is applied only on editions that take it.
    DECLARE @SupportsCompression BIT = CASE WHEN CAST(SERVERPROPERTY('EngineEdition') AS INT) IN (2, 3, 5, 6, 8) THEN 1 ELSE 0 END;

    PRINT 'Starting full backup for database: ' + @DatabaseName;
    PRINT 'Destination file: ' + @BackupPath;

    BEGIN TRY
        IF @SupportsCompression = 1
            BACKUP DATABASE @DatabaseName
            TO DISK = @BackupPath
            WITH
                FORMAT,
                INIT,
                COMPRESSION,
                CHECKSUM,
                STATS = 10,
                NAME = N'BTA_OSG_DB Full Database Backup';
        ELSE
            BACKUP DATABASE @DatabaseName
            TO DISK = @BackupPath
            WITH
                FORMAT,
                INIT,
                CHECKSUM,
                STATS = 10,
                NAME = N'BTA_OSG_DB Full Database Backup';

        PRINT 'Backup completed. Verifying backup integrity...';

        RESTORE VERIFYONLY
        FROM DISK = @BackupPath
        WITH CHECKSUM;

        PRINT 'Verification successful. Backup is valid and ready for disaster recovery.';
    END TRY
    BEGIN CATCH
        DECLARE @Err NVARCHAR(2048) = ERROR_MESSAGE();
        PRINT 'Backup failed: ' + @Err;
        RAISERROR('Backup failed. See the preceding message for the SQL error.', 16, 127) WITH LOG;
    END CATCH;
END;
GO
