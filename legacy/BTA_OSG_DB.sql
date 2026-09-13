-- ============================================================================
-- OFFICE OF THE SECRETARY-GENERAL (OSG) - BTA PARLIAMENT
-- DOCUMENT STATUS TRACKING AND MONITORING SYSTEM
-- Production MS SQL Server Relational Database DDL Schema
-- ============================================================================

IF NOT EXISTS (SELECT * FROM sys.databases WHERE name = 'BTA_OSG_DB')
BEGIN
    CREATE DATABASE BTA_OSG_DB;
END
GO

USE BTA_OSG_DB;
GO

-- 1. Table: tbl_Users (Staff RFID Credentials & RBAC Roles)
IF OBJECT_ID('dbo.tbl_Users', 'U') IS NOT NULL DROP TABLE dbo.tbl_Users;
CREATE TABLE dbo.tbl_Users (
    UserID INT IDENTITY(1,1) PRIMARY KEY,
    RFID_UID VARCHAR(50) NOT NULL UNIQUE,
    FullName VARCHAR(150) NOT NULL,
    UserRole VARCHAR(50) NOT NULL, -- 'Secretary-General', 'OSG Chief', 'System Administrator', 'Administrative Staff'
    IsActive BIT DEFAULT 1,
    CreatedAt DATETIME DEFAULT GETDATE()
);

-- 2. Table: tbl_Documents (Parliamentary Document Master Registry & Landmarks)
IF OBJECT_ID('dbo.tbl_Documents', 'U') IS NOT NULL DROP TABLE dbo.tbl_Documents;
CREATE TABLE dbo.tbl_Documents (
    DocumentID INT IDENTITY(1,1) PRIMARY KEY,
    DocCode VARCHAR(50) NOT NULL UNIQUE, -- e.g. RES-2026-001, BLL-2026-001
    DocType VARCHAR(50) NOT NULL, -- Resolution, Parliament Bill, Committee Report, etc.
    Title NVARCHAR(500) NOT NULL,
    OriginatingOffice NVARCHAR(200) NOT NULL,
    DestinationOffice NVARCHAR(200) NOT NULL,
    CabinetID VARCHAR(50) NOT NULL, -- Storage Landmark: Cabinet
    ShelfNo VARCHAR(50) NOT NULL,   -- Storage Landmark: Shelf
    BoxCode VARCHAR(50) NOT NULL,   -- Storage Landmark: Box
    GDriveURL NVARCHAR(1000) NULL,  -- Cloud PDF Soft Copy Link
    CurrentStatus NVARCHAR(200) DEFAULT 'Received by OSG',
    AssignedStaff NVARCHAR(150) NULL,
    DateReceived DATETIME DEFAULT GETDATE()
);

-- 3. Table: tbl_ActionDirectives (Secretary-General Administrative Orders & Directives)
IF OBJECT_ID('dbo.tbl_ActionDirectives', 'U') IS NOT NULL DROP TABLE dbo.tbl_ActionDirectives;
CREATE TABLE dbo.tbl_ActionDirectives (
    DirectiveID INT IDENTITY(1,1) PRIMARY KEY,
    DocumentID INT NOT NULL FOREIGN KEY REFERENCES dbo.tbl_Documents(DocumentID) ON DELETE CASCADE,
    SGDirective NVARCHAR(200) NOT NULL, -- e.g. For Immediate Action, Referred to Committee on Rules
    AssignedTo NVARCHAR(150) NULL,
    Notes NVARCHAR(1000) NULL,
    LogUser NVARCHAR(150) NOT NULL,
    Timestamp DATETIME DEFAULT GETDATE()
);

-- 4. Table: tbl_AuditTrail (Digital Activity Tracking & Security Logbook)
IF OBJECT_ID('dbo.tbl_AuditTrail', 'U') IS NOT NULL DROP TABLE dbo.tbl_AuditTrail;
CREATE TABLE dbo.tbl_AuditTrail (
    AuditID INT IDENTITY(1,1) PRIMARY KEY,
    UserName NVARCHAR(150) NOT NULL,
    ActionDescription NVARCHAR(1000) NOT NULL,
    Timestamp DATETIME DEFAULT GETDATE()
);

-- Indexes for Fast Query Optimization
CREATE INDEX IX_tbl_Documents_DocCode ON dbo.tbl_Documents(DocCode);
CREATE INDEX IX_tbl_Documents_CabinetID ON dbo.tbl_Documents(CabinetID);
CREATE INDEX IX_tbl_Users_RFID_UID ON dbo.tbl_Users(RFID_UID);
GO

-- Seed Default Staff Credentials
INSERT INTO dbo.tbl_Users (RFID_UID, FullName, UserRole) VALUES
('88A9F321', 'Prof. Ali B. Pangalian', 'Secretary-General'),
('99B1C456', 'Atty. Fatima Z. Rasheed', 'OSG Chief'),
('77C3D987', 'Omire Khalid B. Ebrahim', 'System Administrator'),
('55E5F666', 'Hassim A. Ibrahim', 'Administrative Staff'),
('11A2B3C4', 'CJ Fairoz A. Usop', 'Administrative Staff');
GO
