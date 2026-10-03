USE BTA_OSG_DB;
GO

-- Document Types
IF NOT EXISTS (SELECT 1 FROM dbo.tbl_DocumentTypes WHERE TypeCode = 'RES')
    INSERT INTO dbo.tbl_DocumentTypes (TypeCode, TypeName, Prefix, SortOrder) VALUES ('RES', 'Resolution', 'RES', 1);
IF NOT EXISTS (SELECT 1 FROM dbo.tbl_DocumentTypes WHERE TypeCode = 'BLL')
    INSERT INTO dbo.tbl_DocumentTypes (TypeCode, TypeName, Prefix, SortOrder) VALUES ('BLL', 'Bill', 'BLL', 2);
IF NOT EXISTS (SELECT 1 FROM dbo.tbl_DocumentTypes WHERE TypeCode = 'REP')
    INSERT INTO dbo.tbl_DocumentTypes (TypeCode, TypeName, Prefix, SortOrder) VALUES ('REP', 'Report', 'REP', 3);
IF NOT EXISTS (SELECT 1 FROM dbo.tbl_DocumentTypes WHERE TypeCode = 'EXC')
    INSERT INTO dbo.tbl_DocumentTypes (TypeCode, TypeName, Prefix, SortOrder) VALUES ('EXC', 'Executive Document', 'EXC', 4);
IF NOT EXISTS (SELECT 1 FROM dbo.tbl_DocumentTypes WHERE TypeCode = 'MEM')
    INSERT INTO dbo.tbl_DocumentTypes (TypeCode, TypeName, Prefix, SortOrder) VALUES ('MEM', 'Memorandum', 'MEM', 5);
IF NOT EXISTS (SELECT 1 FROM dbo.tbl_DocumentTypes WHERE TypeCode = 'END')
    INSERT INTO dbo.tbl_DocumentTypes (TypeCode, TypeName, Prefix, SortOrder) VALUES ('END', 'Endorsement', 'END', 6);
GO

-- Document Statuses
IF NOT EXISTS (SELECT 1 FROM dbo.tbl_DocumentStatuses WHERE StatusCode = 'RECEIVED')
    INSERT INTO dbo.tbl_DocumentStatuses (StatusCode, StatusName, SortOrder) VALUES ('RECEIVED', 'Received', 1);
IF NOT EXISTS (SELECT 1 FROM dbo.tbl_DocumentStatuses WHERE StatusCode = 'LOGGED')
    INSERT INTO dbo.tbl_DocumentStatuses (StatusCode, StatusName, SortOrder) VALUES ('LOGGED', 'Logged', 2);
IF NOT EXISTS (SELECT 1 FROM dbo.tbl_DocumentStatuses WHERE StatusCode = 'PENDING_SG')
    INSERT INTO dbo.tbl_DocumentStatuses (StatusCode, StatusName, SortOrder) VALUES ('PENDING_SG', 'Pending SG Review', 3);
IF NOT EXISTS (SELECT 1 FROM dbo.tbl_DocumentStatuses WHERE StatusCode = 'DIRECTIVE_ISSUED')
    INSERT INTO dbo.tbl_DocumentStatuses (StatusCode, StatusName, SortOrder) VALUES ('DIRECTIVE_ISSUED', 'Directive Issued', 4);
IF NOT EXISTS (SELECT 1 FROM dbo.tbl_DocumentStatuses WHERE StatusCode = 'ROUTED')
    INSERT INTO dbo.tbl_DocumentStatuses (StatusCode, StatusName, SortOrder) VALUES ('ROUTED', 'Routed', 5);
IF NOT EXISTS (SELECT 1 FROM dbo.tbl_DocumentStatuses WHERE StatusCode = 'IN_PROGRESS')
    INSERT INTO dbo.tbl_DocumentStatuses (StatusCode, StatusName, SortOrder) VALUES ('IN_PROGRESS', 'In Progress', 6);
IF NOT EXISTS (SELECT 1 FROM dbo.tbl_DocumentStatuses WHERE StatusCode = 'COMPLETED')
    INSERT INTO dbo.tbl_DocumentStatuses (StatusCode, StatusName, SortOrder) VALUES ('COMPLETED', 'Completed', 7);
IF NOT EXISTS (SELECT 1 FROM dbo.tbl_DocumentStatuses WHERE StatusCode = 'ARCHIVED')
    INSERT INTO dbo.tbl_DocumentStatuses (StatusCode, StatusName, SortOrder) VALUES ('ARCHIVED', 'Archived', 8);
IF NOT EXISTS (SELECT 1 FROM dbo.tbl_DocumentStatuses WHERE StatusCode = 'CANCELLED')
    INSERT INTO dbo.tbl_DocumentStatuses (StatusCode, StatusName, SortOrder) VALUES ('CANCELLED', 'Cancelled', 9);
GO

-- Roles
IF NOT EXISTS (SELECT 1 FROM dbo.tbl_Roles WHERE RoleCode = 'SG')
    INSERT INTO dbo.tbl_Roles (RoleCode, RoleName, Description, IsSystemRole) VALUES ('SG', 'Secretary-General', 'Secretary-General', 1);
IF NOT EXISTS (SELECT 1 FROM dbo.tbl_Roles WHERE RoleCode = 'SYSADMIN')
    INSERT INTO dbo.tbl_Roles (RoleCode, RoleName, Description, IsSystemRole) VALUES ('SYSADMIN', 'System Administrator', 'System Administrator', 1);
IF NOT EXISTS (SELECT 1 FROM dbo.tbl_Roles WHERE RoleCode = 'OSG_CHIEF')
    INSERT INTO dbo.tbl_Roles (RoleCode, RoleName, Description, IsSystemRole) VALUES ('OSG_CHIEF', 'OSG Chief', 'OSG Chief', 1);
IF NOT EXISTS (SELECT 1 FROM dbo.tbl_Roles WHERE RoleCode = 'ADMIN_STAFF')
    INSERT INTO dbo.tbl_Roles (RoleCode, RoleName, Description, IsSystemRole) VALUES ('ADMIN_STAFF', 'Administrative Staff', 'Administrative Staff', 0);
GO

-- Permissions
DECLARE @permissions TABLE (Code VARCHAR(50), Name NVARCHAR(100));
INSERT INTO @permissions (Code, Name) VALUES 
('DOC_CREATE', 'Create Documents'),
('DOC_EDIT', 'Edit Documents'),
('DOC_DELETE', 'Delete Documents'),
('DOC_VIEW', 'View Documents'),
('DOC_ROUTE', 'Route Documents'),
('DOC_DIRECTIVE', 'Issue Directives'),
('USER_MANAGE', 'Manage Users'),
('ROLE_MANAGE', 'Manage Roles'),
('SYS_CONFIG', 'System Configuration'),
('REP_VIEW', 'View Reports'),
('AUDIT_VIEW', 'View Audit Logs'),
('STORAGE_MANAGE', 'Manage Storage Locations'),
('DOC_ARCHIVE', 'Archive Documents'),
('RFID_MANAGE', 'Manage RFID Cards'),
('DASHBOARD_VIEW', 'View Dashboard');

INSERT INTO dbo.tbl_Permissions (PermissionCode, PermissionName)
SELECT Code, Name FROM @permissions WHERE NOT EXISTS (SELECT 1 FROM dbo.tbl_Permissions WHERE PermissionCode = Code);
GO

-- Role-Permissions
-- SYSADMIN gets all
INSERT INTO dbo.tbl_RolePermissions (RoleID, PermissionID)
SELECT r.RoleID, p.PermissionID
FROM dbo.tbl_Roles r
CROSS JOIN dbo.tbl_Permissions p
WHERE r.RoleCode = 'SYSADMIN'
AND NOT EXISTS (SELECT 1 FROM dbo.tbl_RolePermissions rp WHERE rp.RoleID = r.RoleID AND rp.PermissionID = p.PermissionID);

-- SG gets DOC_VIEW, DOC_DIRECTIVE, REP_VIEW, DASHBOARD_VIEW
INSERT INTO dbo.tbl_RolePermissions (RoleID, PermissionID)
SELECT r.RoleID, p.PermissionID
FROM dbo.tbl_Roles r
CROSS JOIN dbo.tbl_Permissions p
WHERE r.RoleCode = 'SG' AND p.PermissionCode IN ('DOC_VIEW', 'DOC_DIRECTIVE', 'REP_VIEW', 'DASHBOARD_VIEW')
AND NOT EXISTS (SELECT 1 FROM dbo.tbl_RolePermissions rp WHERE rp.RoleID = r.RoleID AND rp.PermissionID = p.PermissionID);

-- OSG_CHIEF gets most doc and reporting, except sys config/audit/user manage
INSERT INTO dbo.tbl_RolePermissions (RoleID, PermissionID)
SELECT r.RoleID, p.PermissionID
FROM dbo.tbl_Roles r
CROSS JOIN dbo.tbl_Permissions p
WHERE r.RoleCode = 'OSG_CHIEF' AND p.PermissionCode IN ('DOC_CREATE', 'DOC_EDIT', 'DOC_VIEW', 'DOC_ROUTE', 'DOC_DIRECTIVE', 'REP_VIEW', 'STORAGE_MANAGE', 'DOC_ARCHIVE', 'DASHBOARD_VIEW')
AND NOT EXISTS (SELECT 1 FROM dbo.tbl_RolePermissions rp WHERE rp.RoleID = r.RoleID AND rp.PermissionID = p.PermissionID);

-- ADMIN_STAFF gets basic doc tasks
INSERT INTO dbo.tbl_RolePermissions (RoleID, PermissionID)
SELECT r.RoleID, p.PermissionID
FROM dbo.tbl_Roles r
CROSS JOIN dbo.tbl_Permissions p
WHERE r.RoleCode = 'ADMIN_STAFF' AND p.PermissionCode IN ('DOC_CREATE', 'DOC_EDIT', 'DOC_VIEW', 'DOC_ROUTE', 'STORAGE_MANAGE', 'DASHBOARD_VIEW')
AND NOT EXISTS (SELECT 1 FROM dbo.tbl_RolePermissions rp WHERE rp.RoleID = r.RoleID AND rp.PermissionID = p.PermissionID);
GO

-- Users. None. A provisioned server ships with no identity at all: the first run of the
-- desktop app on it opens the claim step, which enrols the real System Administrator and the
-- card that officer will actually tap. Demo profiles seed their own staff in code, never here.
GO

-- Directive Types
DECLARE @ArchivedStatusId INT, @PendingStatusId INT;
SELECT @ArchivedStatusId = StatusID FROM dbo.tbl_DocumentStatuses WHERE StatusCode = 'ARCHIVED';
SELECT @PendingStatusId = StatusID FROM dbo.tbl_DocumentStatuses WHERE StatusCode = 'PENDING_SG';

IF NOT EXISTS (SELECT 1 FROM dbo.tbl_DirectiveTypes WHERE DirectiveCode = 'IMMEDIATE_ACTION')
    INSERT INTO dbo.tbl_DirectiveTypes (DirectiveCode, DirectiveName) VALUES ('IMMEDIATE_ACTION', 'For Immediate Action');
IF NOT EXISTS (SELECT 1 FROM dbo.tbl_DirectiveTypes WHERE DirectiveCode = 'REFER_COMMITTEE')
    INSERT INTO dbo.tbl_DirectiveTypes (DirectiveCode, DirectiveName) VALUES ('REFER_COMMITTEE', 'Referred to Committee');
IF NOT EXISTS (SELECT 1 FROM dbo.tbl_DirectiveTypes WHERE DirectiveCode = 'FORWARD_SPEAKER')
    INSERT INTO dbo.tbl_DirectiveTypes (DirectiveCode, DirectiveName) VALUES ('FORWARD_SPEAKER', 'Forwarded for Speaker Signature');
IF NOT EXISTS (SELECT 1 FROM dbo.tbl_DirectiveTypes WHERE DirectiveCode = 'ADMIN_REVIEW')
    INSERT INTO dbo.tbl_DirectiveTypes (DirectiveCode, DirectiveName) VALUES ('ADMIN_REVIEW', 'Under OSG Administrative Review');
IF NOT EXISTS (SELECT 1 FROM dbo.tbl_DirectiveTypes WHERE DirectiveCode = 'APPROVE_ARCHIVE')
    INSERT INTO dbo.tbl_DirectiveTypes (DirectiveCode, DirectiveName, ResultStatusID) VALUES ('APPROVE_ARCHIVE', 'Approved & Archived', @ArchivedStatusId);
GO

-- Storage Locations
IF NOT EXISTS (SELECT 1 FROM dbo.tbl_StorageLocations WHERE CabinetID = 'CAB-A' AND ShelfNo = 'S-1' AND BoxCode = 'BOX-01')
    INSERT INTO dbo.tbl_StorageLocations (CabinetID, ShelfNo, BoxCode, Description) VALUES ('CAB-A', 'S-1', 'BOX-01', 'Primary filing');
IF NOT EXISTS (SELECT 1 FROM dbo.tbl_StorageLocations WHERE CabinetID = 'CAB-B' AND ShelfNo = 'S-3' AND BoxCode = 'BOX-04')
    INSERT INTO dbo.tbl_StorageLocations (CabinetID, ShelfNo, BoxCode, Description) VALUES ('CAB-B', 'S-3', 'BOX-04', 'Secondary filing');
IF NOT EXISTS (SELECT 1 FROM dbo.tbl_StorageLocations WHERE CabinetID = 'CAB-C' AND ShelfNo = 'S-2' AND BoxCode = 'BOX-02')
    INSERT INTO dbo.tbl_StorageLocations (CabinetID, ShelfNo, BoxCode, Description) VALUES ('CAB-C', 'S-2', 'BOX-02', 'Archive filing');
GO
