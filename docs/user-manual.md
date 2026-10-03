# BTA OSG Document Tracking & Monitoring System - User Manual
**Office of the Secretary-General (OSG), Bangsamoro Transition Authority Parliament**
**Version 2.1 (Production Path A)**

---

## 1. System Overview
The BTA OSG Document Status Tracking & Monitoring System tracks parliamentary documents with real-time status updates, SG action directives, physical landmark vault storage, and audit trails. The sidebar has eight destinations: **Dashboard**, **Data Analytics**, **Document Registry**, **SG Directives**, **Search & Storage**, **User & RFID Admin**, **Audit Trail**, and **Portal Intake**. **Data Analytics** and **Portal Intake** drop out of the sidebar when their switches are off in Station Setup, and that switch takes effect the next time the application starts.

## 2. Authentication (RFID Card Tap)
1. At the login prompt, tap your registered OSG USB RFID badge against the reader.
2. The login window lists the staff enrolled on this station under **Sign in as:**. Pick your name and press **Sign in**, or tap your card on the USB reader: the reader is a keyboard wedge, so a tap submits the card you read in preference to the highlighted entry. A sign in picked from the list is recorded in the audit trail as an on-screen selection, not a badge tap.
3. Once authenticated:
   - System displays your name, assigned role, and access scope.
   - 15-minute inactivity timer begins automatically; after it expires the session ends and you must tap again.
4. Tapping another card immediately switches the authenticated session with audit logging.

## 3. Document Registration
1. Navigate to **Document Registry** on the sidebar.
2. Fill required metadata:
   - **Document Title**: Official title of parliamentary paper.
   - **Document Type**: Regular Communication (COMM), Legislative (LEG), Finance (FIN), or Travel Order (TO). Legacy types (Resolution RES, Parliament Bill BLL, Committee Report REP, Executive Document EXC, Memorandum MEM, Endorsement END) remain valid.
   - **Originating & Destination Offices**.
   - **Physical Storage Landmarks**: Cabinet ID, Shelf No, Box Code.
   - **PDF Attachment**: attach a local .pdf file, or a Google Drive HTTPS link (`drive.google.com` / `docs.google.com`).
   - **Assigned OSG Staff**.
3. Click **Register OSG Document**.
4. System automatically assigns unique sequential code: `PREFIX-yyyy-###` (e.g., `COMM-2026-001`).

## 4. Viewing Document Details & Full History
1. In Registry or Search grid, double-click a document row or select and click **View Details & History**.
2. Tabs available:
   - **Document Overview**: Complete metadata table.
   - **SG Directives Timeline**: Action directives issued by Secretary-General.
   - **Step-by-Step Route & Transmittal Logs**: Inter-office movement and forwarding history.
   - **Storage & Custody Movement History**: Landmark audit trail (Cabinet, Shelf, Box).
3. Actions:
   - **Route**: Forward document to next office with remarks (pick the Action Taken status from the dropdown).
   - **Transfer Storage**: Record vault location change.
   - **Open Link**: Open digital soft-copy in default browser.

## 5. SG Action Directives
*(Restricted to Secretary-General and authorized delegates)*
1. Open **SG Directives** from the sidebar.
2. Select the target document.
3. Select Directive Type (e.g., *For Immediate Action*, *Referred to Committee*, *Forwarded for Speaker Signature*, *Approved & Archived*).
4. Enter directive notes and optional reassignment.
5. Click **Log Action Directive**. Document status and audit trail update immediately.

## 6. Search & Storage Navigation
1. Open **Search & Storage**.
2. Enter keyword (DocCode, Title, Type, Cabinet, Origin, Destination, Staff).
3. Search respects Role-Based Access:
   - Secretary-General, OSG Chief, SysAdmin: View all documents.
   - Administrative Staff: View assigned documents and general intake.

## 7. Portal Intake (when the public portal is enabled)
1. Open **Portal Intake**; the sidebar badge shows how many public submissions are pending.
2. **Refresh Queue** pulls the external submissions; **Import Selected Submission** (Records Section, Secretary-General, System Administrator, or OSG Chief) registers the submission into the official registry and acknowledges it to the portal.
3. **Cloudflare Tunnel** publishes the local portal to a temporary public URL; the URL is copied to your clipboard. **Open OSGPortal** starts the local portal and opens it in your browser.

## 8. Data Analytics
Open **Data Analytics** to see per-section workload, status breakdowns, and deadline pressure for the selected timeframe; charts refresh from the same registry data the grids use.
