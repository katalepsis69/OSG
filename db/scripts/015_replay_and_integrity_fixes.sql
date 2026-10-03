USE BTA_OSG_DB;
GO

-- Migration 015: replay and integrity fixes from the 2026-10-02 audit.
-- 1) tbl_RoutingLogs.ToStatusID becomes nullable: offline replay records custody
--    events whose action is not a status code (revision requests, resubmits,
--    transmittals) and must not invent one. The model already documented this intent.
-- 2) tbl_DocumentAssignments gets the index the per-document snapshot pull probes.
-- 3) FlowDirection is pinned to the two values the application writes.

IF EXISTS (SELECT 1 FROM sys.columns
           WHERE object_id = OBJECT_ID(N'dbo.tbl_RoutingLogs') AND name = 'ToStatusID' AND is_nullable = 0)
BEGIN
    ALTER TABLE dbo.tbl_RoutingLogs ALTER COLUMN ToStatusID INT NULL;
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes
               WHERE name = N'IX_DocumentAssignments_DocumentID'
                 AND object_id = OBJECT_ID(N'dbo.tbl_DocumentAssignments'))
BEGIN
    CREATE INDEX IX_DocumentAssignments_DocumentID
        ON dbo.tbl_DocumentAssignments (DocumentID, AssignmentID DESC);
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.check_constraints
               WHERE name = N'CK_Documents_FlowDirection'
                 AND parent_object_id = OBJECT_ID(N'dbo.tbl_Documents'))
   AND NOT EXISTS (SELECT 1 FROM dbo.tbl_Documents
                   WHERE FlowDirection IS NOT NULL
                     AND FlowDirection NOT IN (N'INCOMING', N'OUTGOING'))
BEGIN
    ALTER TABLE dbo.tbl_Documents ADD CONSTRAINT CK_Documents_FlowDirection
        CHECK (FlowDirection IN (N'INCOMING', N'OUTGOING'));
END
GO
