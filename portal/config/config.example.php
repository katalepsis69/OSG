<?php
// BTA OSG External Intake Portal: Configuration Template
// Copy this file to config.php and set your local environment secrets.
// NEVER commit your live config.php with production credentials to Git.

// Database Configuration (Localhost MySQL 8 / XAMPP)
define('DB_HOST', getenv('PORTAL_DB_HOST') ?: '127.0.0.1');
define('DB_PORT', getenv('PORTAL_DB_PORT') ?: '3306');
define('DB_NAME', getenv('PORTAL_DB_NAME') ?: 'osg');
define('DB_USER', getenv('PORTAL_DB_USER') ?: 'root');
define('DB_PASS', getenv('PORTAL_DB_PASS') !== false ? getenv('PORTAL_DB_PASS') : '');
define('DB_CHARSET', 'utf8mb4');

// Brevo (formerly Sendinblue) REST API Configuration (Port 443 HTTPS)
// Free tier allows up to 300 emails/day.
define('BREVO_API_KEY', getenv('BREVO_API_KEY') ?: '');
define('BREVO_SENDER_EMAIL', 'records@bta-osg.gov.ph');
define('BREVO_SENDER_NAME', 'BTA OSG Records Section');

// Authenticated Bridge API Key (shared secret between VPS and Internal Desktop)
// Verified via hash_equals against the X-Bridge-Key HTTP header.
// Generate a fresh per-install value with 32 random bytes as 64 hex chars (the desktop
// setup wizard has a Generate button) and keep it out of version control. The old
// published demo value must never be used: it gates the whole bridge API.
define('BRIDGE_SECRET_KEY', getenv('BRIDGE_SECRET_KEY') ?: 'REPLACE_WITH_GENERATED_64HEX_KEY');

// On-page OTP display for verify.php. Leave false in every deployed environment; the
// verification code belongs in the email only.
define('ALLOW_DEV_OTP_DISPLAY', false);

// Public Portal Settings
define('PORTAL_BASE_URL', getenv('PORTAL_BASE_URL') ?: 'http://localhost:8085');
define('ORGANIZATION_NAME', 'Bangsamoro Transition Authority - Office of the Secretary-General');
