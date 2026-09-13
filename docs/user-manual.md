# BTA OSG Document Tracking & Monitoring System - User Manual
**Office of the Secretary-General (OSG), Bangsamoro Transition Authority Parliament**
**Version 2.1 (Production Path A)**

---

## 1. System Overview
The BTA OSG Document Status Tracking & Monitoring System tracks parliamentary documents (Bills, Resolutions, Reports, Communications, Memoranda, Endorsements) with real-time status updates, SG action directives, physical landmark vault storage, and audit trails.

## 2. Authentication (RFID Card Tap)
1. At the login prompt, tap your registered OSG USB RFID badge against the reader.
2. If simulator mode is active in development, test buttons (SG, Admin, Staff) are available.
3. Once authenticated:
   - System displays your name, assigned role, and access scope.
   - 15-minute inactivity timer begins automatically.
4. Tapping another card immediately switches the authenticated session with audit logging.

## 3. Document Registration
1. Navigate to **Document Registry** on the sidebar.
2. Fill required metadata:
   - **Document Title**: Official title of parliamentary paper.
   - **Document Type**: Resolution (RES), Parliament Bill (BLL), Committee Report (REP), Executive Document (EXC), Memorandum (MEM), Endorsement (END).
   - **Originating & Destination Offices**.
   - **Physical Storage Landmarks**: Cabinet ID, Shelf No, Box Code.
   - **Google Drive PDF URL**: Must be HTTPS under `drive.google.com` or `docs.google.com`.
   - **Assigned OSG Staff**.
3. Click **Register Parliamentary Document**.
4. System automatically assigns unique sequential code: `PREFIX-yyyy-###` (e.g., `RES-2026-001`).

## 4. Viewing Document Details & Full History
1. In Registry or Search grid, double-click a document row or select and click **View Selected Document Full Specification & History**.
2. Tabs available:
   - **Document Overview**: Complete metadata table.
   - **SG Directives Timeline**: Action directives issued by Secretary-General.
   - **Office Routing Logs**: Inter-office movement and forwarding history.
   - **Physical Storage Movement**: Landmark audit trail (Cabinet, Shelf, Box).
3. Actions:
   - **Route Office Step**: Forward document to next office with remarks.
   - **Transfer Physical Storage**: Record vault location change.
   - **Launch Google Drive PDF**: Open digital soft-copy in default browser.

## 5. SG Action Directives
*(Restricted to Secretary-General and authorized delegates)*
1. Open **SG Directives** from the sidebar.
2. Select the target document.
3. Select Directive Type (e.g., *For Immediate Action*, *Referred to Committee on Rules*, *Forwarded for Speaker Signature*, *Approved & Archived*).
4. Enter directive notes and optional reassignment.
5. Click **Log Action Directive**. Document status and audit trail update immediately.

## 6. Search & Storage Navigation
1. Open **Search & Storage**.
2. Enter keyword (DocCode, Title, Type, Cabinet, Origin, Destination, Staff).
3. Search respects Role-Based Access:
   - Secretary-General, OSG Chief, SysAdmin: View all documents.
   - Administrative Staff: View assigned documents and general intake.
