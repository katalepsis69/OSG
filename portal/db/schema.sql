-- BTA OSG Public Intake Portal & Bridge: MySQL 8 Schema
-- Database: osg (utf8mb4 / utf8mb4_unicode_ci)
-- Privacy invariant: no internal identifier or confidential field ever enters
-- this database. Only public-facing tracking metadata is stored.

CREATE DATABASE IF NOT EXISTS osg
    CHARACTER SET utf8mb4
    COLLATE utf8mb4_unicode_ci;

USE osg;

-- 1. Requesters Table (public citizen / agency contact info)
CREATE TABLE IF NOT EXISTS requesters (
    id INT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
    full_name VARCHAR(150) NOT NULL,
    email VARCHAR(190) NOT NULL UNIQUE,
    phone VARCHAR(20) NOT NULL,
    gender VARCHAR(30) NULL, -- GAD demographic: 'Male', 'Female', 'Prefer not to say'
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 2. Pending Submissions Table (transient pre-OTP submissions)
CREATE TABLE IF NOT EXISTS pending_submissions (
    id CHAR(32) PRIMARY KEY, -- 32-character random token stored in session
    full_name VARCHAR(150) NOT NULL,
    email VARCHAR(190) NOT NULL,
    phone VARCHAR(20) NOT NULL,
    gender VARCHAR(30) NULL, -- GAD demographic, carried to requesters on verification
    document_title VARCHAR(255) NOT NULL,
    category ENUM('COMM', 'LEG', 'FIN', 'TO') NOT NULL,
    status ENUM('pending', 'verified', 'expired') DEFAULT 'pending',
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 3. OTP Codes Table (hashed verification tokens, 10-minute expiry, 5 max attempts)
CREATE TABLE IF NOT EXISTS otp_codes (
    id INT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
    submission_id CHAR(32) NOT NULL,
    code_hash VARCHAR(255) NOT NULL,
    expires_at DATETIME NOT NULL,
    attempts TINYINT DEFAULT 0,
    consumed_at DATETIME NULL,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    KEY idx_submission (submission_id),
    CONSTRAINT fk_otp_submission FOREIGN KEY (submission_id)
        REFERENCES pending_submissions(id) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 4. Atomic Counter Sequences Table
-- Note: Sequences are unique and monotonic, not gapless. Rolled-back transactions
-- burn numbers by design.
CREATE TABLE IF NOT EXISTS counters (
    prefix CHAR(4) NOT NULL,
    yr SMALLINT NOT NULL,
    seq INT UNSIGNED DEFAULT 0,
    PRIMARY KEY (prefix, yr)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- Seed counters for OSG classifications (Year 2026)
INSERT INTO counters (prefix, yr, seq) VALUES
    ('COMM', 2026, 0),
    ('LEG', 2026, 0),
    ('FIN', 2026, 0),
    ('TO', 2026, 0)
ON DUPLICATE KEY UPDATE seq = seq;

-- 5. Public Documents Table
-- Suffix alphabet: 31 characters (ABCDEFGHJKMNPQRSTUVWXYZ23456789), producing
-- 923,521 combinations per sequence number.
CREATE TABLE IF NOT EXISTS documents (
    id INT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
    control_number VARCHAR(24) NOT NULL UNIQUE,
    category ENUM('COMM', 'LEG', 'FIN', 'TO') NOT NULL,
    document_title VARCHAR(255) NOT NULL,
    requester_id INT UNSIGNED NOT NULL,
    public_status ENUM('Received', 'Under Review', 'For Processing', 'Approved', 'Ready for Release', 'Rejected') DEFAULT 'Received',
    imported_at DATETIME NULL, -- Populated by internal bridge import_ack.php
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    KEY idx_imported (imported_at),
    CONSTRAINT fk_doc_requester FOREIGN KEY (requester_id)
        REFERENCES requesters(id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 6. Public Milestones Table
CREATE TABLE IF NOT EXISTS public_milestones (
    id INT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
    document_id INT UNSIGNED NOT NULL,
    public_status ENUM('Received', 'Under Review', 'For Processing', 'Approved', 'Ready for Release', 'Rejected') NOT NULL,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    notified_at DATETIME NULL, -- Populated by hourly Brevo cron worker
    notification_attempts TINYINT UNSIGNED DEFAULT 0,
    last_error VARCHAR(255) NULL,
    KEY idx_notified (notified_at),
    KEY idx_attempts (notification_attempts),
    CONSTRAINT fk_milestone_doc FOREIGN KEY (document_id)
        REFERENCES documents(id) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 7. Public Projection View (used strictly by /track.php for read-only status query)
CREATE OR REPLACE VIEW public_documents AS
SELECT
    control_number,
    category,
    document_title,
    public_status,
    created_at
FROM documents;

-- 8. Dedicated Local User Grants
-- Security Invariant: Exactly one local user 'portal'@'localhost'.
-- User has table-level SELECT on osg.documents and column-level UPDATE on
-- BOTH (public_status, imported_at).
CREATE USER IF NOT EXISTS 'portal'@'localhost' IDENTIFIED BY 'REPLACE_WITH_SECURE_PORTAL_DB_PASSWORD';

GRANT SELECT, INSERT, UPDATE ON osg.pending_submissions TO 'portal'@'localhost';
GRANT SELECT, INSERT, UPDATE, DELETE ON osg.otp_codes TO 'portal'@'localhost';
GRANT SELECT, INSERT, UPDATE ON osg.requesters TO 'portal'@'localhost';
GRANT SELECT, UPDATE ON osg.counters TO 'portal'@'localhost';
GRANT SELECT, INSERT ON osg.documents TO 'portal'@'localhost';
GRANT UPDATE (public_status, imported_at) ON osg.documents TO 'portal'@'localhost';
GRANT SELECT ON osg.public_documents TO 'portal'@'localhost';
GRANT SELECT, INSERT, UPDATE ON osg.public_milestones TO 'portal'@'localhost';

FLUSH PRIVILEGES;
