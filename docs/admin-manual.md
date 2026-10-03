# BTA OSG Document Tracking & Monitoring System - Administrator Manual
**Office of the Secretary-General (OSG), Bangsamoro Transition Authority Parliament**
**Version 2.1 (Production Path A)**

---

## 1. Role & Permissions Architecture (RBAC)
The desktop enforces authorization through two mechanisms working together:

**Section visibility** comes from the seeded role codes (`SG`, `SYSADMIN`, `OSG_CHIEF`,
`RECORDS`, `SECRETARIAT`, `LEGISLATIVE`, `FINANCE`, `TRAVEL`, `ADMIN_STAFF`, db/scripts
007 and 008): manager roles see everything, section desks see their own section plus
documents assigned to them.

**Per-user capability flags** on `tbl_Users` (and mirrored into the desktop cache) decide
what an individual may do at the action sites:

| Capability | Who grants it | Effect when off |
|---|---|---|
| `CanRoute` | User & RFID Admin form | Route refused |
| `CanMove` | User & RFID Admin form | Transfer Storage refused |
| `CanSoftCopy` | User & RFID Admin form | All soft-copy preview, launch, and external-browser paths refused (deny-by-default) |

Fixed action gates that do not depend on flags: directives are issued by the
Secretary-General (or System Administrator) only; portal intake import requires Records
Section, Secretary-General, System Administrator, or OSG Chief; user registration and
station setup require a System Administrator badge. The `tbl_Permissions` rows seeded by
script 007 are the legacy vocabulary (`DOC_CREATE`, `DOC_ROUTE`, ...); the desktop
currently keys on roles and flags, so treat the permission table as reference data until
a release wires it to the action sites.

## 2. Managing Users & RFID Smart Cards
*(System Administrator badge required)*

**First administrator.** A station starts with no accounts at all: provisioning seeds
reference data only, never a user, a role assignment, or a card. Until somebody is claimed,
the registry and the staff list are empty and no card signs in. On a station that has no
System Administrator the app opens the **Claim the first administrator** dialog before the
main window, collecting the officer's full name, their RFID card ID (tapped on the USB reader
or typed into the field, Enter submits), and their desk or department, which defaults to
`ICT / Systems Administration`. That card becomes the only credential that can sign in. The
claim writes to SQL Server in office mode and to the local embedded cache in demo mode, and it
refuses a card already enrolled for a different person, so claiming cannot silently rename and
promote an existing staff member. Closing the dialog without claiming exits the application,
because the station would have no way to sign in.

To administer the remaining staff:
1. Navigate to **User & RFID Admin** tab.
2. Register New User:
   - Enter Full Name, Email, Office.
   - Assign System Role (`Secretary-General`, `OSG Chief`, `System Administrator`, `Administrative Staff`).
   - Scan or enter physical RFID Smart Card UID (Hex).
   - Click **Save User & RFID Smart Card**.
   - Enrolments made while SQL Server is offline are replayed automatically on reconnect,
     so the badge starts working on every workstation after the next sync tick.
3. Locking & Lockout Thresholds:
   - The terminal locks for 5 minutes after 5 consecutive failed card reads on one
     workstation (in-memory, resets on restart).
   - When SQL Server is reachable, each failed tap also increments the card owner's
     `FailedTapCount`; at the threshold the account's `IsLocked` flag is set, and a
     locked account stays locked across restarts (a tap on a locked badge is refused,
     not reset). Unlock from **User & RFID Admin** with **Unlock Selected Account**,
     which clears `IsLocked` and resets `FailedTapCount` to 0 in one step; the SQL
     equivalent is `UPDATE tbl_Users SET IsLocked = 0, FailedTapCount = 0 WHERE UserID = ...`.
   - The login window lists the identities enrolled on this station and signs in the one
     picked, through the same lookup a card tap uses. A USB reader also works: it is a
     keyboard wedge, so its keystrokes are collected by the window and its trailing Enter
     submits the tapped card in preference to the highlighted entry.
   - A sign in chosen from the list is audited as "On-screen badge selection authenticated"
     rather than a badge tap, so the append-only trail never claims a card read that did not
     happen. There is no fixed or published badge number to turn off: the list can only ever
     contain people enrolled on that station.

## 3. Auditing & Compliance
1. Navigate to **Audit Trail** tab.
2. Read-only event stream tracks (codes as written by the application):
   - `LOGIN_SUCCESS`, `LOGOUT`, `SESSION_TIMEOUT`, `SETUP_AUTHORIZED`, `SETUP_DENIED`
   - `DOCUMENT_CREATED`, `DOCUMENT_UPDATED`
   - `DIRECTIVE_ADDED`, `ROUTING_LOG_ADDED`, `STORAGE_LOCATION_CHANGED`
   - `PORTAL_DOCUMENT_IMPORTED`, `PORTAL_BRIDGE_WARNING`, `APP_ERROR`, `OFFLINE_SYNC`
   - Events logged while offline keep their original action type on replay.
3. Audit records are strictly append-only. Modification and deletion are disabled in application code and restricted at database permission level.
