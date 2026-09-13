USE BTA_OSG_DB;
GO

-- Document Indexes
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'UX_Documents_DocCode' AND object_id = OBJECT_ID('dbo.tbl_Documents'))
BEGIN
    CREATE UNIQUE NONCLUSTERED INDEX UX_Documents_DocCode ON dbo.tbl_Documents(DocCode);
END

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_Documents_StatusID' AND object_id = OBJECT_ID('dbo.tbl_Documents'))
BEGIN
    CREATE NONCLUSTERED INDEX IX_Documents_StatusID ON dbo.tbl_Documents(StatusID);
END

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_Documents_TypeID' AND object_id = OBJECT_ID('dbo.tbl_Documents'))
BEGIN
    CREATE NONCLUSTERED INDEX IX_Documents_TypeID ON dbo.tbl_Documents(DocumentTypeID);
END

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_Documents_RegisteredAtUTC' AND object_id = OBJECT_ID('dbo.tbl_Documents'))
BEGIN
    CREATE NONCLUSTERED INDEX IX_Documents_RegisteredAtUTC ON dbo.tbl_Documents(RegisteredAtUTC);
END

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_Documents_Title' AND object_id = OBJECT_ID('dbo.tbl_Documents'))
BEGIN
    CREATE NONCLUSTERED INDEX IX_Documents_Title ON dbo.tbl_Documents(Title);
END

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_Documents_CurrentStorageLocationID' AND object_id = OBJECT_ID('dbo.tbl_Documents'))
BEGIN
    CREATE NONCLUSTERED INDEX IX_Documents_CurrentStorageLocationID ON dbo.tbl_Documents(CurrentStorageLocationID);
END

-- Storage Locations Index
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'UX_StorageLocations_LocationKey' AND object_id = OBJECT_ID('dbo.tbl_StorageLocations'))
BEGIN
    CREATE UNIQUE NONCLUSTERED INDEX UX_StorageLocations_LocationKey ON dbo.tbl_StorageLocations(LocationKey);
END

-- Audit Trail Indexes
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_Audit_EventAtUTC' AND object_id = OBJECT_ID('dbo.tbl_AuditTrail'))
BEGIN
    CREATE NONCLUSTERED INDEX IX_Audit_EventAtUTC ON dbo.tbl_AuditTrail(EventAtUTC);
END

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_Audit_UserID' AND object_id = OBJECT_ID('dbo.tbl_AuditTrail'))
BEGIN
    CREATE NONCLUSTERED INDEX IX_Audit_UserID ON dbo.tbl_AuditTrail(UserID);
END

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_Audit_ActionType' AND object_id = OBJECT_ID('dbo.tbl_AuditTrail'))
BEGIN
    CREATE NONCLUSTERED INDEX IX_Audit_ActionType ON dbo.tbl_AuditTrail(ActionType);
END

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_Audit_EntityType_EntityID' AND object_id = OBJECT_ID('dbo.tbl_AuditTrail'))
BEGIN
    CREATE NONCLUSTERED INDEX IX_Audit_EntityType_EntityID ON dbo.tbl_AuditTrail(EntityType, EntityID);
END

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_Audit_DocumentCode' AND object_id = OBJECT_ID('dbo.tbl_AuditTrail'))
BEGIN
    CREATE NONCLUSTERED INDEX IX_Audit_DocumentCode ON dbo.tbl_AuditTrail(DocumentCode);
END
GO
