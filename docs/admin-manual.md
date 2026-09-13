# BTA OSG Document Tracking & Monitoring System - Administrator Manual
**Office of the Secretary-General (OSG), Bangsamoro Transition Authority Parliament**
**Version 2.1 (Production Path A)**

---

## 1. Role & Permissions Architecture (RBAC)
The system enforces strict RBAC across 4 baseline roles:

| Permission | SG | SYSADMIN | OSG_CHIEF | ADMIN_STAFF |
|---|---|---|---|---|
| `DOCUMENT_VIEW_ALL` | Yes | Yes | Yes | No |
| `DOCUMENT_VIEW_ASSIGNED` | Yes | Yes | Yes | Yes |
| `DOCUMENT_CREATE` | Yes | Yes | Yes | Yes |
| `DOCUMENT_UPDATE` | Yes | Yes | Yes | Assigned only |
| `DOCUMENT_DELETE_OR_DISABLE` | No | Yes | No | No |
| `DIRECTIVE_ISSUE` | Yes | No | No | No |
| `ROUTING_CREATE` | Yes | Yes | Yes | No by default |
| `STORAGE_UPDATE` | Yes | Yes | Yes | Yes |
| `PDF_OPEN` | Yes | Yes | Yes | Yes |
| `AUDIT_VIEW` | Yes | Yes | Yes | No |
| `USER_MANAGE` | No | Yes | No | No |
| `RFID_MANAGE` | No | Yes | No | No |
| `ROLE_MANAGE` | No | Yes | No | No |
| `REFERENCE_MANAGE` | No | Yes | No | No |
| `REPORT_EXPORT` | Yes | Yes | Yes | No |

## 2. Managing Users & RFID Smart Cards
*(SYSADMIN role required)*
1. Navigate to **User & RFID Admin** tab.
2. Register New User:
   - Enter Full Name, Email, Office.
   - Assign System Role (`Secretary-General`, `OSG Chief`, `System Administrator`, `Administrative Staff`).
   - Scan or enter physical RFID Smart Card UID (Hex).
   - Click **Save User & RFID Smart Card**.
3. Locking & Lockout Thresholds:
   - User accounts lock automatically after 5 consecutive failed tap attempts.
   - Unlock user by clearing `IsLocked` flag and resetting `FailedTapCount` to 0.
   - Unknown card taps are recorded separately in `tbl_RfidFailedAttempts`.

## 3. Auditing & Compliance
1. Navigate to **Audit Trail** tab.
2. Read-only event stream tracks:
   - `LOGIN_SUCCESS`, `LOGIN_FAILURE`, `LOGOUT`, `SESSION_TIMEOUT`
   - `DOCUMENT_CREATED`, `DOCUMENT_UPDATED`, `DOCUMENT_STATUS_CHANGED`
   - `DIRECTIVE_ADDED`, `ROUTING_LOG_ADDED`, `STORAGE_LOCATION_CHANGED`
   - `PDF_LINK_OPENED`
3. Audit records are strictly append-only. Modification and deletion are disabled in application code and restricted at database permission level.
