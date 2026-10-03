USE BTA_OSG_DB;
GO

-- 007 used to seed one bootstrap System Administrator whose card value was published in the
-- repository, so any reader could walk up to any installed station and tap into it. New
-- servers no longer get that identity. This retires it on servers that already do, but only
-- once a claimed administrator exists, because revoking the last administrator would lock
-- every workstation out of its own office.
DECLARE @BootstrapUserId INT = (SELECT UserID FROM dbo.tbl_Users WHERE Username = 'oebrahim');

DECLARE @ClaimedAdministrators INT = (
    SELECT COUNT(1)
    FROM dbo.tbl_Users u
    JOIN dbo.tbl_UserRoles ur ON ur.UserID = u.UserID
    JOIN dbo.tbl_Roles r ON r.RoleID = ur.RoleID
    WHERE r.RoleCode = 'SYSADMIN' AND ur.IsActive = 1 AND u.IsActive = 1
      AND u.UserID <> ISNULL(@BootstrapUserId, -1));

IF @BootstrapUserId IS NOT NULL AND @ClaimedAdministrators > 0
BEGIN
    UPDATE dbo.tbl_RfidCards
    SET IsActive = 0,
        RevokedAtUTC = SYSUTCDATETIME(),
        RevokedByUserID = NULL,
        RevocationReason = 'Published bootstrap card retired'
    WHERE UserID = @BootstrapUserId AND IsActive = 1;

    UPDATE dbo.tbl_UserRoles SET IsActive = 0 WHERE UserID = @BootstrapUserId AND IsActive = 1;

    -- Rows stay in place: documents, routing logs, and audit events name this UserID, and the
    -- trail is append-only, so access is removed and history is kept.
    UPDATE dbo.tbl_Users SET IsActive = 0 WHERE UserID = @BootstrapUserId;
END
GO
