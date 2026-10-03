# Office Flow and Setup

A plain guide for the office that runs this system: what each part does, how they connect, and how to set it up. The technical detail lives in SETUP_GUIDE.md and deployment.md.

## What the system does

It tracks office documents from arrival to release. An officer signs in with a badge, registers a document, sends it to another desk, and every step is recorded. Only this program writes to the office records.

## The parts

1. The program officers work in. Register, route, store, print, audit.
2. The office records. One set, held on the server PC.
3. The public portal. Where a citizen submits a request and checks it later. A separate program with its own separate records.
4. Portal Intake. A tab inside the program, where an officer reads portal submissions and brings them in.
5. Data Analytics. Charts made from the documents the system already holds.
6. Routing slip. One page that travels with the folder.

The portal and the charts are optional. Switch either off in Station Setup and nothing else changes.

## The daily document flow

1. Receive and register. When a paper document arrives, the Records desk enters its details into Document Registry. The program assigns an official tracking code (like COMM-2026-001).
2. Print the routing slip. Print the one-page slip and clip it to the physical folder.
3. Move the folder. The physical folder is carried to the next desk or office.
4. Update the step. The receiving officer signs in, opens the document, records their action or directive, and routes it forward or files it in the vault.

## How they connect


        Citizen on a phone, anywhere
                    |
                    v   public link
        +------------------------+
        |  Public portal         |
        |  Own separate records  |
        +------------------------+
                    |
                    |   the program asks for the
                    |   waiting list of submissions
                    v
        +------------------------+
        |  The program officers  |
        |  work in. They sign in |
        |  with their badges     |
        +------------------------+
                    |
                    v
        +------------------------+
        |  The office records,   |
        |  on the server PC      |
        +------------------------+

        One way back: when an officer changes a
        status, the program tells the portal so
        the citizen's tracking page updates.
        Nothing else goes backwards.


A submission waits in the portal until a Records officer opens Portal Intake and presses Import Selected Submission. Until that click, the office knows nothing about it.

## Why the portal cannot change the office records

The portal is a different program holding its own records, and it is never given the address or the password for the office records. Only three things pass between them: the program asks for the waiting list, reports that a submission was brought in, and sends back one status word.

Bringing a submission in leaves it where it is. The portal keeps its own copy for the citizen's tracking page, so removing a document here does not remove it there.

## First time, on the server PC

1. Install SQL Server Express on that PC. The default name it picks is fine.
2. Copy the program folder across. If an offline record file came along with it, delete it.
3. Run BTA_OSG_DocumentTracking.exe. A Connect window opens.
4. Click Server address and other settings.
5. Type the server address as localhost\SQLEXPRESS. Leave Port empty.
6. Type the shared office password. This becomes the password for the office records, so choose one and keep it safe.
7. Press Connect. The first attempt fails because that password is not set up yet. That is expected.
8. Set up this server appears. Press it and allow the prompt. It builds the office records and checks the connection again.
9. A claim window opens. Type the administering officer's name, tap the badge or type its card ID, adjust the desk if needed, then press Claim administrator. That card becomes the administrator badge for this station.
10. Sign in and enrol the other staff under User & RFID Admin.

## Every other PC in the office

1. Copy the same program folder across. Delete any offline record file in the copy first.
2. Run the program, type the same shared password, press Connect.
3. The address is found and remembered. Sign in with an enrolled badge.

If the computer does not find the server automatically: click Server address and other settings. Type the server PC's computer name or local network IP address (for example OFFICE-SERVER\SQLEXPRESS or 192.168.1.50\SQLEXPRESS), then press Connect.

There is no second setup to run. Set up this server only appears on a PC that holds its own copy of SQL Server, and pressing it there would build a second set of records nobody is looking at.

## What does not come along with the copy

- The saved address and password. Each desk is asked for the shared password once.
- The offline record file. It holds one PC's own notes, so never copy it.
- The portal. It stays on the server PC.
- A printer, since routing slips print from the PC they are made on.
- The document preview helper. Some PCs lack a small Microsoft piece it needs. Without it the Preview tab says so and Open in External Browser still works.

What does arrive ready: the program needs nothing else installed to run, and every enrolled staff member reaches each desk, because the staff list lives in the office records.

## Does the office need internet?

No. The desktop program and the office records communicate entirely over your local office network (network cables or office Wi-Fi). If the internet goes down, officers can still register, route, search, and audit documents without interruption.

Internet is only needed for two optional features:
1. The public citizen portal and its Cloudflare tunnel.
2. Opening soft-copy PDF documents hosted on Google Drive links.

## Start with the program, add the portal later

Run the office on the program alone first. Officers can register, route, store, print, and audit walk-in documents without the portal, and that is the whole office workflow. Turn the portal on in a second pass, once its email and its public address belong to the office.

To switch it on later: open Station Setup, tick Enable Citizen Web Portal Intake, restart the program, and the Portal Intake tab appears.

## Running the portal

Nothing extra needs installing. The portal carries its own pieces and the program starts them when the portal is switched on. This is why the portal folder is large.

Two things to settle before it faces the public.

Email has to work, or submissions cannot finish. A citizen is sent a six-digit code and must type it in before the request counts. If that email does not arrive, nobody can submit and Portal Intake stays empty, because a submission only joins the waiting list after its code is entered. The copy on this machine sends from a personal Gmail account set up for testing, so the office's own mail account has to be put in before the public uses it. That change belongs to whoever looks after the office's computers, not to the officers.

The public address changes every time. The portal listens only to the PC it runs on, so access from outside goes through a free service that hands out a new address on each start. Until the office has a permanent address, do not print the link on a form, a flyer, or a sign.

## How a citizen uses it on a phone

1. On the server PC, sign in, open Portal Intake, press Open OSGPortal to show the portal is alive, then press Cloudflare Tunnel. After a few seconds the program shows the public link and copies it, so you can paste it into a message. Press the same button later to see or copy that link again.
2. The citizen opens the link on a phone. It needs its own internet, so mobile data is enough and the phone does not need to be near the office.
3. They fill in name, email, phone, and the document details, then submit. The portal emails a six-digit code.
4. They type that code in. Only now the portal gives them a control number and sends a receipt.
5. Tell them to keep the receipt or write down the control number. Without it they cannot track the request.

The link dies when the PC sleeps, when the program closes, or when someone stops the tunnel. Leave the server PC awake and signed in during office hours.

## Daily routine: morning and evening

Morning:
Turn on the server PC first and ensure it stays awake. Other desks cannot connect if the server PC is turned off or in sleep mode.

Evening and backups:
The server PC holds the only master copy of all office records.
At the end of the day or week, copy the latest SQL Server backup file onto a separate USB flash drive, and keep that drive in a secure drawer. Whoever manages the office computers can schedule this backup task to run automatically.

## No card reader yet

The login window lists every staff member enrolled in the system, each with their role. Pick a name and press Sign in. The audit trail records it as an on-screen selection rather than a card tap, so the record stays honest.

## Practise mode, no server at all

On the Connect window, click Run without a server (demo mode). The program keeps its records in a file on that PC. Use it to practise the flow before the office sees it.

To move that PC onto the office records later, press Connect to the Office Server (Station Setup) at the bottom of the login screen. Because the station already has settings saved, it asks for the administrator badge first, and only an administrator gets through. Type the server address and the shared password, press Connect, and the window confirms the station is in office mode without needing a restart.

## Quick help for common hiccups

- A workstation says Connection failed: Check that the server PC is turned on and awake. Check that both computers are plugged into the same office router or connected to the same office Wi-Fi.
- An officer forgot their badge: Select their name from the Sign in as list on the login window and press Sign in.
- Card tap says terminal locked: Five failed card taps lock that computer for 5 minutes. Wait 5 minutes or restart the program. An administrator can also clear the lock under User & RFID Admin.

## Test these before go-live

1. Register a document and check the code it is given.
2. Route it to another desk, sign in there, and confirm it appears.
3. Print the routing slip.
4. Open Audit Trail and confirm those three actions are listed.
5. On a second PC, confirm the same document appears on the next sync, about ten seconds later.
6. With the portal on, submit a request from a phone, complete the email code, import it in Portal Intake, and confirm it appears in the registry.
