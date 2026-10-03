# BTA OSG Public Intake Portal & Bridge: Deployment Runbook

**Office of the Secretary-General (OSG)**  
**Bangsamoro Transition Authority Parliament**  
Document Version: 1.0 (Production Release)  
Target Environment: Oracle Cloud Infrastructure (Always-Free Tier) + Ubuntu Linux + PHP 8 + MySQL 8 + Nginx

---

## 1. System Architecture & Core Invariants

The BTA OSG Document Tracking ecosystem consists of two synchronized environments connected by a thin authenticated HTTPS bridge:

1. **Internal Office System (System of Record):**
   - Platform: VB.NET WinForms (.NET 10) on SQL Server 2022.
   - Network: Restricted to OSG LAN; air-gapped from direct public exposure.
   - Authentication: RFID Smart Card Badge authentication with 6 RBAC roles.
   - Audit Trail: Immutable append-only audit trail logging all actions.

2. **External Intake Portal (Public Submissions & Status Tracking):**
   - Platform: PHP 8 + MySQL 8 on an isolated cloud VPS.
   - Ingress: Public HTTPS (Port 443).
   - Functions: Citizen registration, email OTP verification, control number assignment, read-only tracking.
   - Email: Dispatched over HTTPS port 443 via the Brevo REST API (native cURL; bypasses VPS outbound SMTP port blocks).

3. **Core Isolation Invariant:**
   - With `PortalEnabled = false` (or if the cloud VPS is completely offline), the internal office system runs 100% normally with zero interruption, freeze, or blocking. No portal failure may ever compromise internal operations.
   - Privacy Invariant: No internal identifier, confidential note, or internal employee credential ever enters the public portal database.

---

## 2. Cloud VPS Provisioning (Oracle Cloud Always-Free Tier)

### Step 2.1: Instance Selection
Log in to the Oracle Cloud Infrastructure (OCI) Console and create a Compute Instance:
- **Image:** Canonical Ubuntu 22.04 LTS or 24.04 LTS (Minimal or Server).
- **Shape Option A (Recommended):** `VM.Standard.A1.Flex` (Ampere ARM, 2 to 4 OCPUs, 12 to 24 GB RAM: Always Free).
- **Shape Option B (Fallback):** `VM.Standard.E2.1.Micro` (AMD x86_64, 1 OCPU, 1 GB RAM: Always Free).

### Step 2.2: Cloud Gotcha 1: OCI A1 Capacity Retry
In high-demand cloud regions (e.g., Singapore, Tokyo, Sydney), attempting to launch an A1 Ampere instance may occasionally return `Out of capacity for shape VM.Standard.A1.Flex`.
- **Immediate Resolution:** Switch shape to `VM.Standard.E2.1.Micro`. The micro shape is always available and possesses more than enough capacity to comfortably host Nginx, PHP 8.2-FPM, and MySQL 8 for the public intake portal.
- **Automated Retry Option:** Use an OCI CLI loop script or Terraform provider with exponential backoff if ARM architecture is preferred.

### Step 2.3: Cloud Gotcha 2: OCI Dual-Firewall Trap
OCI instances feature two separate layers of packet filtering. Opening ports in the web console alone is insufficient:
1. **Cloud Layer (VCN Security List / NSG):**
   - Navigate to **Networking > Virtual Cloud Networks > Default Security List for VCN**.
   - Add Ingress Rules:
     - Source: `0.0.0.0/0`, IP Protocol: TCP, Destination Port Range: `80`
     - Source: `0.0.0.0/0`, IP Protocol: TCP, Destination Port Range: `443`
2. **Host OS Layer (Ubuntu iptables):**
   - OCI Ubuntu images include default `iptables` rules that drop non-SSH ingress packets.
   - Run the following commands on the VPS:
     ```bash
     sudo iptables -I INPUT 6 -m state --state NEW -p tcp --dport 80 -j ACCEPT
     sudo iptables -I INPUT 6 -m state --state NEW -p tcp --dport 443 -j ACCEPT
     sudo netfilter-persistent save
     ```
   *(Note: `oracle-setup.sh` automates these iptables adjustments).*

---

## 3. Server Setup & Software Installation

### Step 3.1: Run Host Setup Script
SSH into your provisioned cloud instance:
```bash
ssh -i /path/to/private_key ubuntu@YOUR_VPS_PUBLIC_IP
```

Upload or clone the repository to the server, then execute the deployment script:
```bash
chmod +x portal/deploy/oracle-setup.sh
sudo ./portal/deploy/oracle-setup.sh
```

The script installs:
- Nginx Web Server
- PHP 8.2 with FPM, cURL, MySQL, and mbstring extensions
- MySQL Server 8.0 (configured to bind strictly to `127.0.0.1`)
- Certbot with Nginx plugin

### Step 3.2: Initialize the MySQL Database Schema
Log into MySQL as root and apply the schema:
```bash
sudo mysql < portal/db/schema.sql
```

The schema creates:
- Database `osg` with `utf8mb4` encoding.
- Dedicated local user `'portal'@'localhost'`.
- Strict granular table and column privileges per the defense specification.
- Monotonic sequence counters seeded for Year 2026 (`COMM`, `LEG`, `FIN`, `TO`).
- Read-only projection view `public_documents` for `/track.php`.

Set a strong password for the `'portal'@'localhost'` MySQL user:
```sql
ALTER USER 'portal'@'localhost' IDENTIFIED BY 'YOUR_GENERATED_SECURE_PASSWORD';
FLUSH PRIVILEGES;
```

---

## 4. Web Application Deployment & Configuration

### Step 4.1: Deploy Files to Web Root
```bash
sudo mkdir -p /var/www/osg-portal
sudo cp -r portal/* /var/www/osg-portal/
sudo chown -R www-data:www-data /var/www/osg-portal
sudo chmod -R 750 /var/www/osg-portal
```

### Step 4.2: Configure Secrets in `config.php`
Copy the configuration template:
```bash
sudo cp /var/www/osg-portal/config/config.example.php /var/www/osg-portal/config/config.php
sudo chown www-data:www-data /var/www/osg-portal/config/config.php
sudo chmod 600 /var/www/osg-portal/config/config.php
```

The cron jobs write their log under `/var/log/osg-portal/`, which www-data cannot
create by itself, so make the directory now:
```bash
sudo mkdir -p /var/log/osg-portal
sudo chown www-data:www-data /var/log/osg-portal
```

Edit `/var/www/osg-portal/config/config.php`:
```php
// Database Configuration
define('DB_HOST', '127.0.0.1');
define('DB_PORT', '3306');
define('DB_NAME', 'osg');
define('DB_USER', 'portal');
define('DB_PASS', 'YOUR_GENERATED_SECURE_PASSWORD');

// Brevo HTTP API Configuration (Port 443 HTTPS)
define('BREVO_API_KEY', 'xkeysib-YOUR_ACTUAL_BREVO_KEY');
define('BREVO_SENDER_EMAIL', 'records@bta-osg.gov.ph');
define('BREVO_SENDER_NAME', 'BTA OSG Records Section');

// Authenticated Bridge API Key
define('BRIDGE_SECRET_KEY', 'YOUR_RANDOM_64_CHARACTER_HEX_BRIDGE_KEY');

// Portal Settings
define('PORTAL_BASE_URL', 'https://portal.bta-osg.gov.ph');
```

### Step 4.3: Bridge Key Hygiene
Generate a cryptographically secure 64-character hex key:
```bash
openssl rand -hex 32
```
- Place this key into `BRIDGE_SECRET_KEY` in `/var/www/osg-portal/config/config.php`.
- Place the identical key into the internal office `appsettings.Production.json` under `Portal.BridgeKey`.
- **Hygiene Rule:** Never commit `config.php` or production secrets to Git.

### Step 4.4: Cloud Gotcha 3: Brevo Sender Email Verification
Free cloud VPS providers block outbound SMTP ports (25, 465, 587). The portal bypasses this restriction by communicating directly with Brevo's REST API over HTTPS port 443 using native cURL.
- **Requirement:** Before Brevo permits email dispatch, you must verify your sender email address.
- In the Brevo dashboard, navigate to **Senders, Domains & Dedicated IPs > Senders**.
- Add and verify the sender address (e.g., `records@bta-osg.gov.ph` or your registered administrator email).
- If using an unverified sender address, the Brevo API returns HTTP 400 with `Sender not allowed`.

---

## 5. Nginx & Let's Encrypt SSL Configuration

### Step 5.1: Install Nginx Configuration
```bash
sudo cp /var/www/osg-portal/deploy/nginx-osg-portal.conf /etc/nginx/sites-available/osg-portal.conf
sudo ln -sf /etc/nginx/sites-available/osg-portal.conf /etc/nginx/sites-enabled/
sudo rm -f /etc/nginx/sites-enabled/default
sudo nginx -t
sudo systemctl reload nginx
```

### Step 5.2: Obtain SSL Certificate
Ensure your DNS A record points `portal.bta-osg.gov.ph` to your VPS public IP, then issue:
```bash
sudo certbot --nginx -d portal.bta-osg.gov.ph --non-interactive --agree-tos -m admin@bta-osg.gov.ph
```

---

## 6. Scheduled Cron Automation & Backups

### Step 6.1: Hourly Milestone Notification Worker
Install the crontab for the `www-data` web user:
```bash
sudo crontab -u www-data /var/www/osg-portal/cron/crontab.txt
```

Verify the crontab:
```bash
sudo crontab -u www-data -l
```

### Step 6.2: Automated Daily Backups
Two prerequisites before the schedule works:
1. The backup script reads its MySQL credentials from `/etc/osg-portal-backup.cnf`
   (chmod 600, owned by root). Create it once:
   ```bash
   sudo tee /etc/osg-portal-backup.cnf >/dev/null <<'EOF'
   [client]
   user=portal
   password=THE_PORTAL_USER_PASSWORD
   EOF
   sudo chown root:root /etc/osg-portal-backup.cnf
   sudo chmod 600 /etc/osg-portal-backup.cnf
   ```
2. The cron log directory from Step 4.2 (`/var/log/osg-portal`) exists, so the
   job's output redirect cannot abort it.

Then install the schedule in ROOT's crontab (the script reads the root-owned
credential file, so it cannot run as www-data):
```bash
sudo crontab -e
```
Add line:
```cron
30 2 * * * /bin/bash /var/www/osg-portal/deploy/backup-portal.sh >> /var/log/osg-portal/backup.log 2>&1
```
Install the backup schedule exactly once: the www-data crontab template
(`portal/cron/crontab.txt`) deliberately carries no backup line, so a copy of both
files does not produce two dumps a night.

---

## 7. Live Defense Demonstration Script (Walkthrough)

Follow these exact steps during panel defense and evaluation:

### Step 1: Public Document Registration
1. Navigate in a browser to `https://portal.bta-osg.gov.ph` (or local test host).
2. Enter submitter details:
   - Full Name: `Amina S. Macacua`
   - Official Email: `amina.macacua@example.com`
   - Mobile: `09171234567`
   - Classification: `Legal Opinions & Legislative Documents (LEG)`
   - Title: `Formal Request for Legal Opinion on Parliamentary Bill No. 88`
   - Check the **Data Privacy Act (RA 10173)** consent box.
3. Click **Proceed to Email Verification**.

### Step 2: OTP Verification & Receipt Issuance
1. The portal redirects to `/verify.php?token=...`.
2. Check email for the 6-digit verification code.
3. Enter the 6-digit numeric code on the screen and submit.
4. The system transactionally:
   - Increments the atomic counter for `LEG-2026`.
   - Generates a 4-character suffix from the 31-character alphabet.
   - Assigns control number: `LEG-2026-0001-K9X2`.
   - Sends an official confirmation receipt email.
   - Renders the printable official receipt card on screen.

### Step 3: Office Staff RFID Login & Queue Inspection
1. Launch the internal desktop application (`BTA_OSG_DocumentTracking.exe`).
2. Tap the Records Section officer's enrolled RFID card (the badge the administrator enrolled for that officer on the **User & RFID Admin** tab; on a station that has no accounts yet, the first run opens **Claim the first administrator** before the main window, and that claimed card is the only credential that signs in until more staff are enrolled).
3. Click on the 7th sidebar tab: **Portal Intake**.
4. The status indicator displays: `Portal Bridge: Connected`.
5. Click **Refresh Queue**. The newly submitted document appears in the queue grid.

### Step 4: Ingestion into Internal Registry
1. Select the row for `LEG-2026-0001-K9X2`.
2. Click **Import Selected Submission**.
3. Confirm the prompt dialog.
4. The system:
   - Generates the internal document code (e.g. `LEG-2026-001`).
   - Stores the linked `ExternalControlNumber = LEG-2026-0001-K9X2`.
   - Auto-routes to the **Legislative Section**.
   - Sends an import acknowledgment to the portal bridge via `import_ack.php`.
   - Logs an audit event in `tbl_AuditTrail`.
5. The submission drops out of the portal intake queue.

### Step 5: Internal Action & Workflow Status Transition
1. Switch user via RFID badge to the Legislative Section officer's enrolled card, or the Secretary-General's enrolled card.
2. Open Document Detail for `LEG-2026-001`.
3. Issue a directive or route the document:
   - Select **Approve Document** (or **Request Revision**).
4. The internal system automatically and asynchronously executes `PushPublicStatusAsync`:
   - Calls bridge `POST /api/status.php` with `public_status: Approved`.
   - A public milestone is logged with `notified_at = NULL`.

### Step 6: Citizen Live Status Tracking
1. Switch back to the public browser at `https://portal.bta-osg.gov.ph/track.php`.
2. Enter `LEG-2026-0001-K9X2` into the search box.
3. The page displays:
   - Document Title and Category.
   - Current Status Pill: `APPROVED` (in green).
   - Full chronological timeline showing `Received` and `Approved` milestones with timestamps.

### Step 7: Hourly Milestone Email Dispatch
1. Run the milestone cron worker:
   ```bash
   php /var/www/osg-portal/cron/send_milestones.php
   ```
2. The worker picks up the unnotified milestone and dispatches an update email via Brevo HTTP API.
3. The citizen receives an email confirming their document status transitioned to `Approved`.

### Step 8: Honest Neutral Not-Found Test
1. In `track.php`, search for a non-existent control number: `LEG-2026-9999-ZZZZ`.
2. The system responds with an honest, neutral alert:
   > *"No registered document was found with control number LEG-2026-9999-ZZZZ. Please verify the code on your official receipt and check for typographical errors."*
3. Demonstrates proper error handling and absence of generic error slop.
