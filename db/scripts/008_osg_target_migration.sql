USE BTA_OSG_DB;
GO

-- Target Document Types
IF NOT EXISTS (SELECT 1 FROM dbo.tbl_DocumentTypes WHERE TypeCode = 'REG_COMM')
    INSERT INTO dbo.tbl_DocumentTypes (TypeCode, TypeName, Prefix, SortOrder) VALUES ('REG_COMM', 'Regular Communication', 'COMM', 1);
IF NOT EXISTS (SELECT 1 FROM dbo.tbl_DocumentTypes WHERE TypeCode = 'LEG')
    INSERT INTO dbo.tbl_DocumentTypes (TypeCode, TypeName, Prefix, SortOrder) VALUES ('LEG', 'Legislative', 'LEG', 2);
IF NOT EXISTS (SELECT 1 FROM dbo.tbl_DocumentTypes WHERE TypeCode = 'FIN')
    INSERT INTO dbo.tbl_DocumentTypes (TypeCode, TypeName, Prefix, SortOrder) VALUES ('FIN', 'Finance', 'FIN', 3);
IF NOT EXISTS (SELECT 1 FROM dbo.tbl_DocumentTypes WHERE TypeCode = 'TRAVEL')
    INSERT INTO dbo.tbl_DocumentTypes (TypeCode, TypeName, Prefix, SortOrder) VALUES ('TRAVEL', 'Travel Order', 'TO', 4);
GO

-- Target Document Statuses
IF NOT EXISTS (SELECT 1 FROM dbo.tbl_DocumentStatuses WHERE StatusCode = 'RECEIVED')
    INSERT INTO dbo.tbl_DocumentStatuses (StatusCode, StatusName, SortOrder) VALUES ('RECEIVED', 'Received', 1);
IF NOT EXISTS (SELECT 1 FROM dbo.tbl_DocumentStatuses WHERE StatusCode = 'FOR_REVIEW')
    INSERT INTO dbo.tbl_DocumentStatuses (StatusCode, StatusName, SortOrder) VALUES ('FOR_REVIEW', 'For Review', 2);
IF NOT EXISTS (SELECT 1 FROM dbo.tbl_DocumentStatuses WHERE StatusCode = 'FOR_REVISION')
    INSERT INTO dbo.tbl_DocumentStatuses (StatusCode, StatusName, SortOrder) VALUES ('FOR_REVISION', 'For Revision', 3);
IF NOT EXISTS (SELECT 1 FROM dbo.tbl_DocumentStatuses WHERE StatusCode = 'APPROVED')
    INSERT INTO dbo.tbl_DocumentStatuses (StatusCode, StatusName, SortOrder) VALUES ('APPROVED', 'Approved', 4);
IF NOT EXISTS (SELECT 1 FROM dbo.tbl_DocumentStatuses WHERE StatusCode = 'RELEASED')
    INSERT INTO dbo.tbl_DocumentStatuses (StatusCode, StatusName, SortOrder) VALUES ('RELEASED', 'Released', 5);
IF NOT EXISTS (SELECT 1 FROM dbo.tbl_DocumentStatuses WHERE StatusCode = 'FILED')
    INSERT INTO dbo.tbl_DocumentStatuses (StatusCode, StatusName, SortOrder) VALUES ('FILED', 'Filed', 6);
GO

-- Column Extensions on tbl_Documents
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'dbo.tbl_Documents') AND name = 'FlowDirection')
    ALTER TABLE dbo.tbl_Documents ADD FlowDirection VARCHAR(10) NOT NULL CONSTRAINT DF_tbl_Documents_FlowDirection DEFAULT 'INCOMING';

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'dbo.tbl_Documents') AND name = 'AssignedSection')
    ALTER TABLE dbo.tbl_Documents ADD AssignedSection NVARCHAR(100) NULL;

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'dbo.tbl_Documents') AND name = 'TargetDeadlineUTC')
    ALTER TABLE dbo.tbl_Documents ADD TargetDeadlineUTC DATETIME2 NULL;

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'dbo.tbl_Documents') AND name = 'RevisionPunchlist')
    ALTER TABLE dbo.tbl_Documents ADD RevisionPunchlist NVARCHAR(MAX) NULL;

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'dbo.tbl_Documents') AND name = 'LastActionTaken')
    ALTER TABLE dbo.tbl_Documents ADD LastActionTaken NVARCHAR(255) NULL;
GO

-- Target Section Roles
IF NOT EXISTS (SELECT 1 FROM dbo.tbl_Roles WHERE RoleCode = 'RECORDS')
    INSERT INTO dbo.tbl_Roles (RoleCode, RoleName, Description, IsSystemRole) VALUES ('RECORDS', 'Records Section', 'Records intake and scanning', 1);
IF NOT EXISTS (SELECT 1 FROM dbo.tbl_Roles WHERE RoleCode = 'SECRETARIAT')
    INSERT INTO dbo.tbl_Roles (RoleCode, RoleName, Description, IsSystemRole) VALUES ('SECRETARIAT', 'Secretariat', 'Regular communications review', 1);
IF NOT EXISTS (SELECT 1 FROM dbo.tbl_Roles WHERE RoleCode = 'LEGISLATIVE')
    INSERT INTO dbo.tbl_Roles (RoleCode, RoleName, Description, IsSystemRole) VALUES ('LEGISLATIVE', 'Legislative Section', 'Parliament legislative docs', 1);
IF NOT EXISTS (SELECT 1 FROM dbo.tbl_Roles WHERE RoleCode = 'FINANCE')
    INSERT INTO dbo.tbl_Roles (RoleCode, RoleName, Description, IsSystemRole) VALUES ('FINANCE', 'Finance Section', 'Budgetary and financial communications', 1);
IF NOT EXISTS (SELECT 1 FROM dbo.tbl_Roles WHERE RoleCode = 'TRAVEL')
    INSERT INTO dbo.tbl_Roles (RoleCode, RoleName, Description, IsSystemRole) VALUES ('TRAVEL', 'Travel Section', 'Travel order administration', 1);
GO
