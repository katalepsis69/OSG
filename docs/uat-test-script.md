# BTA OSG Document Tracking & Monitoring System - UAT Test Script
**Office of the Secretary-General (OSG), Bangsamoro Transition Authority Parliament**
**Version 2.1 (Production Path A)**

Fill in the Pass/Fail column while executing; a row marked PASS before the step has been
run is not a test result.

**Preparation (do this before row UAT-01):** the application ships with no accounts and no
known card numbers. Launch it, complete the **Claim the first administrator** dialog naming
the officer who will administer the station, sign in with that card, then use the **User &
RFID Admin** tab to enrol a badge for every role this script exercises: Secretary-General, OSG
Chief, System Administrator (a second officer if the office has one), and one officer for each
section desk the script touches, such as Finance and Records. Record the card ID you enrolled
for each officer beside their row below; every step that says "enrolled card" means the
physical badge you issued in this preparation step, not a pre-installed identity.

| Test ID | Scenario | Steps | Expected Result | Pass/Fail |
|---|---|---|---|---|
| **UAT-01** | RFID Smart Card Login | On first run, confirm no card signs in until the claim completes, then claim the administrator and sign in with that card. Enrol the Secretary-General officer's badge from the Admin tab and tap it. | Before the claim the app has no identity to authenticate and the **Claim the first administrator** dialog opens ahead of the main window; closing it without claiming exits the application. After the claim, tapping each enrolled card signs that officer in and the header displays their name, with `[SECRETARY-GENERAL]: GLOBAL ACCESS` for the Secretary-General badge. | |
| **UAT-02** | Invalid Card Rejection | Tap unregistered card UID `AA11BB22` five times | Access denied warning displayed each time; on the fifth tap the terminal locks for 5 minutes and (with SQL reachable) the account's `FailedTapCount` climbs in `tbl_Users`. | |
| **UAT-03** | Auto-Switch User Session | Tap the enrolled Finance Section officer's card while the Secretary-General session is active | Session switches immediately to that officer with the `[FINANCE SECTION]` scope; audit logs the session switch. | |
| **UAT-04** | Document Registration | Fill form with Title: "Draft Resolution on Regional Water Security", Type: "Regular Communication" | Auto-generates sequential code `COMM-2026-###`; status set to `FOR_REVIEW` and routed to the Secretariat. | |
| **UAT-05** | Sequential Code Minting | Register two documents of the same type in a row | Codes increment (`COMM-2026-001`, `COMM-2026-002`); SQL Server's unique index on `tbl_Documents.DocCode` backstops duplicates across workstations. | |
| **UAT-06** | SG Directive Issuance | Under SG session, select document and apply directive "For Immediate Action" | Directive timeline updated; document status unchanged (immediate action carries no status outcome); audit logged with `DIRECTIVE_ADDED`. | |
| **UAT-07** | Inter-Office Routing | Route document from OSG to "Committee on Rules" | New entry in Routing Logs; origin/destination updated; status change and routing log land together. | |
| **UAT-08** | Storage Physical Transfer | Change landmark to `CAB-B / S-2 / BOX-05` with reason | Movement history logged; physical location string updated. | |
| **UAT-09** | Soft-Copy Security Validation | Enter URL `http://untrusted-site.com/doc.pdf` | Blocked with security warning (requires HTTPS and `drive.google.com`); a non-PDF local file is also refused by "Open Link". | |
| **UAT-10** | RBAC Visibility Filter | Tap an enrolled section desk officer's card (for example the Finance Section officer enrolled during preparation) | The desk sees only documents assigned to them or their section, not the full archive. | |
| **UAT-11** | Append-Only Audit Trail | Sign in with the claimed System Administrator's card and view the Audit tab | All actions (logins, registrations, directives, routings, the `ADMIN_CLAIMED` entry written by the claim, `LOGOUT`, `SESSION_TIMEOUT`) appear read-only; no update or delete path exists. | |
| **UAT-12** | Inactivity Session Timeout | Log in, then leave the workstation idle for 15 minutes | Session automatically ends with a `SESSION_TIMEOUT` audit entry; header shows logged out; tap required to resume. | |
| **UAT-13** | Offline Enrolment Replay | Disable SQL connectivity, register a new staff badge, restore connectivity | Badge works on the enrolling workstation immediately and on other workstations after the next sync tick; the user survives the snapshot refresh. | |
| **UAT-14** | Restart Keeps Offline Work | Disable SQL connectivity, register a document, close and reopen the app | The offline document reappears after restart (the XML cache reloads) and replays to SQL Server when connectivity returns, keeping the code it was registered under unless the code collides. | |
