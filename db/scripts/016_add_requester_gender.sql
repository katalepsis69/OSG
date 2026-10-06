-- 016: Portal gender capture as a first-class column.
-- The portal collects requester gender (GAD demographic data) and the offline
-- cache carries it in RequesterGender, but tbl_Documents only kept it as free
-- text inside Remarks, so it was unqueryable in the office registry. This adds
-- the column the import path writes and the analytics reads.
-- Idempotent: safe to re-run on any server generation.

IF COL_LENGTH('dbo.tbl_Documents', 'RequesterGender') IS NULL
BEGIN
    ALTER TABLE dbo.tbl_Documents ADD RequesterGender NVARCHAR(20) NULL;
END
GO
