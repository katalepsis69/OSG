# BTA OSG Document Tracking & Monitoring System - UAT Test Script
**Office of the Secretary-General (OSG), Bangsamoro Transition Authority Parliament**
**Version 2.1 (Production Path A)**

---

| Test ID | Scenario | Steps | Expected Result | Pass/Fail |
|---|---|---|---|---|
| **UAT-01** | RFID Smart Card Login | Tap card `88A9F321` (Secretary-General) | Header displays `Prof. Ali B. Pangalian [SECRETARY-GENERAL] — GLOBAL ACCESS`. | PASS |
| **UAT-02** | Invalid Card Rejection | Tap unregistered card UID `AA11BB22` | Access denied warning displayed; attempt logged in `tbl_RfidFailedAttempts`. | PASS |
| **UAT-03** | Auto-Switch User Session | Tap Staff card `55E5F666` while SG session active | Session switches immediately to `Hassim A. Ibrahim [ADMINISTRATIVE STAFF]`; audit logs session switch. | PASS |
| **UAT-04** | Document Registration | Fill form with Title: "Draft Resolution on Regional Water Security", Type: "Resolution" | Auto-generates sequential code `RES-2026-###`; status set to `LOGGED`. | PASS |
| **UAT-05** | Duplicate Code Prevention | Attempt to register document with identical code | Blocked by database unique constraint with friendly user prompt. | PASS |
| **UAT-06** | SG Directive Issuance | Under SG session, select document and apply directive "For Immediate Action" | Directive timeline updated; document status reflects directive; audit logged. | PASS |
| **UAT-07** | Inter-Office Routing | Route document from OSG to "Committee on Rules" | New entry in Routing Logs; origin/destination updated. | PASS |
| **UAT-08** | Storage Physical Transfer | Change landmark to `CAB-B / S-2 / BOX-05` with reason | Movement history logged; physical location string updated. | PASS |
| **UAT-09** | Google Drive Security Validation | Enter URL `http://untrusted-site.com/doc.pdf` | Blocked with security warning (requires HTTPS and `drive.google.com`). | PASS |
| **UAT-10** | RBAC Visibility Filter | Login as Administrative Staff (`55E5F666`) | Staff sees only documents assigned to them or their office, not full archive. | PASS |
| **UAT-11** | Append-Only Audit Trail | Login as SysAdmin (`77C3D987`) and view Audit tab | All actions (logins, registrations, directives, routings) appear read-only. | PASS |
| **UAT-12** | Inactivity Session Timeout | Idle system for 15 minutes | Session automatically terminates; return to login screen. | PASS |
