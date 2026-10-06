-- ============================================================================
-- purge_seed_leftovers.sql  (MANUAL, never provisioned)
-- ============================================================================
-- Final seed policy (2026-10-03): the office server never carries seed or test
-- data. Older test runs wrote their probe rows straight into BTA_OSG_DB before
-- the suite moved to its throwaway catalog; this script removes those
-- leftovers. It is deliberately NOT part of the 001-015 provisioning chain and
-- never runs on its own.
--
-- HOW TO RUN (server machine, sqlcmd):
--   sqlcmd -S "localhost\SQLEXPRESS" -E -C -i db\scripts\purge_seed_leftovers.sql
--
-- STEP 1 below only LISTS what matched. Review it first. STEP 2 performs the
-- deletes; keep it commented out until the list looks right, because documents
-- an operator registered for real can share patterns with probe titles.
-- The audit trail is append-only: rows naming purged probe documents are the
-- one sanctioned exception, and only because they name test artifacts. Deleting
-- a mid-chain audit row breaks the 014 hash-chain verification for every
-- subsequent entry, and the chain cannot be repaired by the app, which only
-- fills NULL hashes. Run a backup first (see backup_sqlserver.sql).
-- ============================================================================

USE BTA_OSG_DB;
GO

-- Probe document titles the integration suite and self-check mint.
-- Match by title, not by DocCode: probe codes are stamped per run.
DECLARE @probeTitlePatterns TABLE (Pattern NVARCHAR(200));
INSERT INTO @probeTitlePatterns (Pattern) VALUES
    ('E2E connected registration probe%'),
    ('E2E mirror read probe%'),
    ('E2E offline replay probe%'),
    ('E2E storage move probe%'),
    ('Clean replay probe%'),
    ('Collision probe%'),
    ('Conflict probe%'),
    ('Guard probe%'),
    ('Null status probe%'),
    ('Self-check isolation probe%'),
    ('Self-check sequential mint probe%'),
    ('Self-Check Workflow Verification');

-- ============================================================================
-- STEP 1: REVIEW ONLY. Run the script as-is, read the lists, then decide.
-- ============================================================================
SELECT d.DocumentID, d.DocCode, d.Title, d.RegisteredAtUTC
FROM dbo.tbl_Documents d
JOIN @probeTitlePatterns p ON d.Title LIKE p.Pattern
ORDER BY d.DocumentID;

-- Probe staff: the suite's enrolled officer and any replayed self-check badges.
SELECT u.UserID, u.Username, u.FullName, u.IsActive
FROM dbo.tbl_Users u
WHERE u.Username = 'probeofficer'
   OR u.FullName LIKE 'Self-Check %'
   OR u.FullName LIKE 'E2E Replay Staff%';

-- ============================================================================
-- STEP 2: DELETE. Uncomment the block below only after reviewing STEP 1.
-- ============================================================================
/*
DECLARE @probeDocs TABLE (DocumentID INT, DocCode VARCHAR(30));
INSERT INTO @probeDocs (DocumentID, DocCode)
SELECT d.DocumentID, d.DocCode FROM dbo.tbl_Documents d
JOIN @probeTitlePatterns p ON d.Title LIKE p.Pattern;

BEGIN TRANSACTION;

DELETE al FROM dbo.tbl_AuditTrail al
JOIN @probeDocs t ON al.DocumentCode = t.DocCode;

DELETE rl FROM dbo.tbl_RoutingLogs rl
JOIN @probeDocs t ON rl.DocumentID = t.DocumentID;
DELETE da FROM dbo.tbl_DocumentAssignments da
JOIN @probeDocs t ON da.DocumentID = t.DocumentID;
DELETE mv FROM dbo.tbl_DocumentMovements mv
JOIN @probeDocs t ON mv.DocumentID = t.DocumentID;
DELETE ad FROM dbo.tbl_ActionDirectives ad
JOIN @probeDocs t ON ad.DocumentID = t.DocumentID;
DELETE d FROM dbo.tbl_Documents d
JOIN @probeDocs t ON d.DocumentID = t.DocumentID;

DELETE c FROM dbo.tbl_RfidCards c
JOIN dbo.tbl_Users u ON c.UserID = u.UserID
WHERE u.Username = 'probeofficer'
   OR u.FullName LIKE 'Self-Check %'
   OR u.FullName LIKE 'E2E Replay Staff%';
DELETE r FROM dbo.tbl_UserRoles r
JOIN dbo.tbl_Users u ON r.UserID = u.UserID
WHERE u.Username = 'probeofficer'
   OR u.FullName LIKE 'Self-Check %'
   OR u.FullName LIKE 'E2E Replay Staff%';
DELETE f FROM dbo.tbl_RfidFailedAttempts f
JOIN dbo.tbl_Users u ON f.UserID = u.UserID
WHERE u.Username = 'probeofficer'
   OR u.FullName LIKE 'Self-Check %'
   OR u.FullName LIKE 'E2E Replay Staff%';

DELETE u FROM dbo.tbl_Users u
WHERE u.Username = 'probeofficer'
   OR u.FullName LIKE 'Self-Check %'
   OR u.FullName LIKE 'E2E Replay Staff%';

COMMIT TRANSACTION;

SELECT N'Purged probe documents' AS Result, COUNT(*) AS RowsInDocuments FROM dbo.tbl_Documents;
*/
GO
