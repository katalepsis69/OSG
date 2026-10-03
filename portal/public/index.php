<?php
// BTA OSG External Intake Portal: Public Submission Page
// Bangsamoro Transition Authority - Office of the Secretary-General
// Collects public submissions, validates input, dispatches 6-digit email OTP.

declare(strict_types=1);

require_once __DIR__ . '/../lib/db.php';
require_once __DIR__ . '/../lib/mail.php';
require_once __DIR__ . '/../lib/ratelimit.php';
require_once __DIR__ . '/../lib/session.php';
portal_session_start();

$errors = [];
$formData = [
    'full_name' => '',
    'email' => '',
    'phone' => '',
    'gender' => 'Prefer not to say',
    'category' => 'COMM',
    'document_title' => '',
    'consent' => ''
];

$allowedCategories = [
    'COMM' => 'General Communications & Correspondence (COMM)',
    'LEG'  => 'Legal Opinions & Legislative Documents (LEG)',
    'FIN'  => 'Financial & Budgetary Submissions (FIN)',
    'TO'   => 'Special Orders & Travel Directives (TO)'
];

$allowedGenders = [
    'Male' => 'Male',
    'Female' => 'Female',
    'Prefer not to say' => 'Prefer not to say'
];

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    // 0. CSRF: the intake form posts a per-session token; cross-site forgeries rejected.
    portal_csrf_verify($_POST['csrf_token'] ?? null);

    // 0b. Per-IP throttle on the OTP-dispatching endpoint: the per-email limiter below
    // rotates with the victim's address, so the inbox-bombing and quota-exhaustion
    // defence needs a bucket keyed on the sender's own IP.
    if (!osg_rate_limit('submit', 8, 600)) {
        $errors[] = 'Too many submissions from your IP address. Please try again later.';
    }

    // 1. Anti-bot honeypot check
    if (!empty($_POST['website_hp'] ?? '')) {
        // Silent blackhole for bots
        sleep(2);
        header('Location: index.php?status=submitted');
        exit;
    }

    // 2. Extract and sanitize input
    $formData['full_name'] = trim((string)($_POST['full_name'] ?? ''));
    $formData['gender'] = trim((string)($_POST['gender'] ?? 'Prefer not to say'));
    $formData['email'] = trim((string)($_POST['email'] ?? ''));
    $formData['phone'] = trim((string)($_POST['phone'] ?? ''));
    $formData['category'] = trim((string)($_POST['category'] ?? 'COMM'));
    $formData['document_title'] = trim((string)($_POST['document_title'] ?? ''));
    $formData['consent'] = isset($_POST['privacy_consent']) ? '1' : '';

    // 3. Validation
    if ($formData['full_name'] === '' || mb_strlen($formData['full_name']) > 150) {
        $errors[] = 'Please enter your full name (maximum 150 characters).';
    }

    if (!array_key_exists($formData['gender'], $allowedGenders)) {
        $errors[] = 'Please select a valid gender option.';
    }

    if (!filter_var($formData['email'], FILTER_VALIDATE_EMAIL) || mb_strlen($formData['email']) > 190) {
        $errors[] = 'Please enter a valid email address.';
    }

    // Standard Philippine phone number validation (e.g. 0917 123 4567 or +63 917 123 4567)
    $cleanPhone = preg_replace('/[\s\-]/', '', $formData['phone']);
    if (str_starts_with($cleanPhone, '639') && strlen($cleanPhone) === 12) {
        $cleanPhone = '+' . $cleanPhone;
    }
    if (!preg_match('/^(09\d{9}|\+639\d{9})$/', $cleanPhone)) {
        $errors[] = 'Please enter a valid Philippine mobile number (e.g. 0917 123 4567 or +63 917 123 4567).';
    }

    if (!array_key_exists($formData['category'], $allowedCategories)) {
        $errors[] = 'Please select a valid document category.';
    }

    if ($formData['document_title'] === '' || mb_strlen($formData['document_title']) > 255) {
        $errors[] = 'Please enter the document title or subject (maximum 255 characters).';
    }

    if ($formData['consent'] !== '1') {
        $errors[] = 'You must agree to the Data Privacy Act (RA 10173) consent statement.';
    }

    // 4. Rate limiting: Check recent unverified pending submissions from this email (max 5 per hour)
    if (empty($errors)) {
        try {
            $pdo = get_db();
            $rateStmt = $pdo->prepare(
                'SELECT COUNT(*) FROM pending_submissions WHERE email = ? AND created_at > DATE_SUB(NOW(), INTERVAL 1 HOUR)'
            );
            $rateStmt->execute([$formData['email']]);
            $recentCount = (int)$rateStmt->fetchColumn();

            if ($recentCount >= 5) {
                $errors[] = 'Too many verification attempts for this email address. Please try again after 1 hour.';
            }
        } catch (Throwable $ex) {
            error_log('[DB Error] Rate limit query failed: ' . $ex->getMessage());
            $errors[] = 'A system error occurred. Please try again later.';
        }
    }

    // 5. Create pending submission and dispatch OTP
    if (empty($errors)) {
        try {
            $pdo = get_db();
            $pdo->beginTransaction();

            $submissionId = bin2hex(random_bytes(16)); // 32-char hex token
            $otpCode = (string)random_int(100000, 999999);
            $otpHash = password_hash($otpCode, PASSWORD_DEFAULT);

            // Insert pending submission
            $subStmt = $pdo->prepare(
                'INSERT INTO pending_submissions (id, full_name, email, phone, gender, document_title, category, status, created_at)
                 VALUES (?, ?, ?, ?, ?, ?, ?, "pending", NOW())'
            );
            $subStmt->execute([
                $submissionId,
                $formData['full_name'],
                $formData['email'],
                $cleanPhone,
                $formData['gender'],
                $formData['document_title'],
                $formData['category']
            ]);

            // Insert OTP record valid for 10 minutes
            $otpStmt = $pdo->prepare(
                'INSERT INTO otp_codes (submission_id, code_hash, expires_at, attempts, created_at)
                 VALUES (?, ?, DATE_ADD(NOW(), INTERVAL 10 MINUTE), 0, NOW())'
            );
            $otpStmt->execute([$submissionId, $otpHash]);

            $pdo->commit();

            // Dispatch verification email
            // The OTP lives in the email body only: subject lines propagate into mail
            // provider logs, desktop search indexes, and notification previews.
            $subject = 'Verify your email to complete your BTA OSG document submission';
            $html = '
            <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; color: #1e293b; line-height: 1.6;">
                <div style="background-color: #0f2a4a; padding: 20px; color: #ffffff; text-align: center;">
                    <h2 style="margin: 0; font-size: 1.25rem;">Bangsamoro Transition Authority</h2>
                    <p style="margin: 4px 0 0 0; font-size: 0.875rem; color: #cbd5e1;">Office of the Secretary-General</p>
                </div>
                <div style="padding: 24px; border: 1px solid #cbd5e1; background-color: #ffffff;">
                    <p>Dear <strong>' . htmlspecialchars($formData['full_name'], ENT_QUOTES, 'UTF-8') . '</strong>,</p>
                    <p>Thank you for submitting your document to the BTA OSG. To complete your submission and receive your official control number, please verify your email address using the one-time code below:</p>
                    <div style="margin: 24px 0; text-align: center;">
                        <span style="display: inline-block; font-family: monospace; font-size: 2rem; font-weight: bold; letter-spacing: 0.5rem; background-color: #f1f5f9; padding: 12px 24px; border: 1px solid #cbd5e1; border-radius: 6px; color: #0f2a4a;">
                            ' . htmlspecialchars($otpCode, ENT_QUOTES, 'UTF-8') . '
                        </span>
                    </div>
                    <p style="font-size: 0.875rem; color: #64748b;">This verification code is valid for <strong>10 minutes</strong>. Do not share this code with anyone.</p>
                    <hr style="border: none; border-top: 1px solid #e2e8f0; margin: 20px 0;">
                    <p style="font-size: 0.8125rem; color: #64748b;">If you did not initiate this document submission, you can safely ignore this email.</p>
                </div>
            </div>';

            $mailRes = send_brevo_mail($formData['email'], $formData['full_name'], $subject, $html);
            if (!$mailRes['success']) {
                // Details go to the server log only: provider diagnostics on the public
                // page leak the mail provider's error text to anonymous visitors.
                $mailErr = $mailRes['error'] ?? 'Unknown mail error';
                error_log('[Mail Warning] Verification email dispatch returned failure: ' . $mailErr);
                $errors[] = 'We could not deliver a verification email right now. Please check the address and try again, or contact the Records Section.';
            } else {
                $_SESSION['pending_token'] = $submissionId;
                // Dev-only OTP display: stored solely when the explicit config flag is on
                // (local thesis demos). Tunnel traffic appears as 127.0.0.1, so a REMOTE_ADDR
                // check cannot distinguish local from tunneled and is deliberately not used.
                // The temp file exists only under the same flag and is unlinked on verify.
                if (defined('ALLOW_DEV_OTP_DISPLAY') && ALLOW_DEV_OTP_DISPLAY === true) {
                    $_SESSION['dev_last_otp'] = $otpCode;
                    $devOtpPath = sys_get_temp_dir() . DIRECTORY_SEPARATOR . 'bta_dev_otp_' . $submissionId;
                    @file_put_contents($devOtpPath, $otpCode);
                    @chmod($devOtpPath, 0600);
                }
                header('Location: verify.php?token=' . urlencode($submissionId));
                exit;
            }

        } catch (Throwable $ex) {
            if ($pdo && $pdo->inTransaction()) {
                $pdo->rollBack();
            }
            error_log('[Submission Error] ' . $ex->getMessage());
            $errors[] = 'A server error occurred while processing your submission. Please try again.';
        }
    }
}
$activePage = 'index';
?>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Submit Document - OSGPortal Public Intake</title>
    <link rel="icon" type="image/x-icon" href="favicon.ico">
    <link rel="stylesheet" href="css/portal.css?v=<?= filemtime(__DIR__ . '/css/portal.css') ?>">
</head>
<body>
    <header class="portal-header">
        <div class="portal-header-container">
            <img class="portal-brand-seal" src="assets/bta_logo.png" alt="Bangsamoro Parliament Official Seal" width="56" height="56">
            <div class="portal-brand">
                <span class="portal-brand-sub">Republic of the Philippines &bull; Bangsamoro Autonomous Region in Muslim Mindanao</span>
                <span class="portal-brand-title">Bangsamoro Transition Authority</span>
                <span class="portal-brand-subtitle">Office of the Secretary-General &bull; OSGPortal Public Intake</span>
            </div>
            <nav class="portal-nav" aria-label="Public Portal Navigation">
                <a href="index.php" class="portal-nav-link <?= ($activePage === 'index') ? 'active' : '' ?>">Submit Document</a>
                <a href="track.php" class="portal-nav-link <?= ($activePage === 'track') ? 'active' : '' ?>">Track Document</a>
            </nav>
        </div>
    </header>

    <main class="portal-main">
        <div class="portal-layout-wrapper">
            <div class="portal-layout-main">
                <section class="portal-card">
                    <div class="portal-stepper" aria-label="Intake Journey Progress">
                        <div class="portal-step-item active">
                            <span class="portal-step-num">1</span>
                            <span>Submit Details</span>
                        </div>
                        <div class="portal-step-item">
                            <span class="portal-step-num">2</span>
                            <span>Verify Email (OTP)</span>
                        </div>
                        <div class="portal-step-item">
                            <span class="portal-step-num">3</span>
                            <span>Track Online &amp; Receipt</span>
                        </div>
                        <div class="portal-step-item">
                            <span class="portal-step-num">4</span>
                            <span>Digital Document Release</span>
                        </div>
                    </div>

                    <div class="portal-card-header">
                        <h1 class="portal-card-title">Public Document Registration</h1>
                        <p class="portal-card-subtitle">
                            Register external correspondence, legislative measures, or official requests for chain-of-custody tracking by the BTA OSG Records Section.
                        </p>
                    </div>

            <?php if (($_GET['notice'] ?? '') === 'expired'): ?>
                <div class="portal-alert portal-alert-warning" role="alert">
                    <span class="portal-alert-title">Session Expired:</span>
                    Your verification session is no longer valid. Please submit your details again to receive a new one-time code.
                </div>
            <?php endif; ?>

            <?php if (!empty($errors)): ?>
                <div class="portal-alert portal-alert-danger" role="alert">
                    <span class="portal-alert-title">Please review the following errors:</span>
                    <ul style="margin-left: 1.25rem; margin-top: 0.25rem;">
                        <?php foreach ($errors as $error): ?>
                            <li><?= htmlspecialchars($error, ENT_QUOTES, 'UTF-8') ?></li>
                        <?php endforeach; ?>
                    </ul>
                </div>
            <?php endif; ?>

            <form action="index.php" method="POST" novalidate
                  onsubmit="if (this.dataset.busy) { return false; } this.dataset.busy = '1'; var b = this.querySelector('button[type=submit]'); if (b) { b.disabled = true; b.textContent = 'Submitting...'; }">
                <!-- CSRF token -->
                <input type="hidden" name="csrf_token" value="<?= htmlspecialchars(portal_csrf_token(), ENT_QUOTES, 'UTF-8') ?>">

                <!-- Honeypot Field -->
                <div class="portal-hp" aria-hidden="true">
                    <label for="website_hp">Leave this field blank</label>
                    <input type="text" id="website_hp" name="website_hp" tabindex="-1" autocomplete="off">
                </div>

                <div class="portal-form-grid">
                    <!-- Row 1: Full Name & Gender -->
                    <div class="portal-col-6">
                        <div class="portal-form-group">
                            <label class="portal-label" for="full_name">
                                Full Name / Authorized Representative <span class="required">*</span>
                            </label>
                            <input type="text" id="full_name" name="full_name" class="portal-input"
                                   maxlength="150" required autocomplete="name"
                                   value="<?= htmlspecialchars($formData['full_name'], ENT_QUOTES, 'UTF-8') ?>"
                                   placeholder="e.g. Juan Dela Cruz">
                            <p class="portal-help-text">Individual submitting the document or representing an organization.</p>
                        </div>
                    </div>

                    <div class="portal-col-6">
                        <div class="portal-form-group">
                            <label class="portal-label" for="gender">
                                Requester Gender (GAD Demographic) <span class="required">*</span>
                            </label>
                            <select id="gender" name="gender" class="portal-select" required>
                                <?php foreach ($allowedGenders as $val => $lbl): ?>
                                    <option value="<?= $val ?>" <?= $formData['gender'] === $val ? 'selected' : '' ?>>
                                        <?= htmlspecialchars($lbl, ENT_QUOTES, 'UTF-8') ?>
                                    </option>
                                <?php endforeach; ?>
                            </select>
                            <p class="portal-help-text">Used for Gender and Development (GAD) compliance and statistics.</p>
                        </div>
                    </div>

                    <!-- Row 2: Official Email & Mobile Phone -->
                    <div class="portal-col-6">
                        <div class="portal-form-group">
                            <label class="portal-label" for="email">
                                Official Email Address <span class="required">*</span>
                            </label>
                            <input type="email" id="email" name="email" class="portal-input"
                                   maxlength="190" required autocomplete="email"
                                   value="<?= htmlspecialchars($formData['email'], ENT_QUOTES, 'UTF-8') ?>"
                                   placeholder="juan.delacruz@example.com">
                            <p class="portal-help-text">A one-time verification code and official tracking receipt will be sent here.</p>
                        </div>
                    </div>

                    <div class="portal-col-6">
                        <div class="portal-form-group">
                            <label class="portal-label" for="phone">
                                Mobile Phone Number <span class="required">*</span>
                            </label>
                            <input type="tel" id="phone" name="phone" class="portal-input"
                                   maxlength="20" required autocomplete="tel"
                                   value="<?= htmlspecialchars($formData['phone'], ENT_QUOTES, 'UTF-8') ?>"
                                   placeholder="0917 123 4567">
                            <p class="portal-help-text">Philippine mobile format (e.g. 0917 123 4567 or +63 917 123 4567).</p>
                        </div>
                    </div>

                    <!-- Row 3: Classification Category (Full Width) -->
                    <div class="portal-col-12">
                        <div class="portal-form-group">
                            <label class="portal-label" for="category">
                                Document Classification / Category <span class="required">*</span>
                            </label>
                            <select id="category" name="category" class="portal-select" required>
                                <?php foreach ($allowedCategories as $code => $label): ?>
                                    <option value="<?= $code ?>" <?= $formData['category'] === $code ? 'selected' : '' ?>>
                                        <?= htmlspecialchars($label, ENT_QUOTES, 'UTF-8') ?>
                                    </option>
                                <?php endforeach; ?>
                            </select>
                        </div>
                    </div>

                    <!-- Row 4: Document Title / Subject (Full Width) -->
                    <div class="portal-col-12">
                        <div class="portal-form-group">
                            <label class="portal-label" for="document_title">
                                Document Title / Subject <span class="required">*</span>
                            </label>
                            <input type="text" id="document_title" name="document_title" class="portal-input"
                                   maxlength="255" required
                                   value="<?= htmlspecialchars($formData['document_title'], ENT_QUOTES, 'UTF-8') ?>"
                                   placeholder="e.g. Formal Request for Legal Opinion on Parliamentary Bill No. 45">
                            <p class="portal-help-text">Brief, descriptive title or subject matter of the document.</p>
                        </div>
                    </div>

                    <!-- Row 4: Digital Intake Notice Callout -->
                    <div class="portal-col-12">
                        <div class="portal-alert portal-alert-info">
                            <div class="portal-alert-title">Digital Submission Notice</div>
                            <div>
                                Documents registered through this portal are directly recorded in the BTA OSG digital intake queue. Upon completing email verification, you will receive an official External Control Number to track your document's status in real time.
                            </div>
                        </div>
                    </div>

                    <!-- Row 5: Data Privacy Consent Checkbox -->
                    <div class="portal-col-12">
                        <div class="portal-checkbox-group">
                            <input type="checkbox" id="privacy_consent" name="privacy_consent" value="1"
                                   <?= $formData['consent'] === '1' ? 'checked' : '' ?> required>
                            <label for="privacy_consent" class="portal-checkbox-label">
                                <strong>Data Privacy Consent (RA 10173):</strong> I consent to the collection and processing of my contact information and document details solely for the purposes of public document tracking, verification, and status monitoring by the Bangsamoro Transition Authority - Office of the Secretary-General in compliance with the Data Privacy Act of 2012.
                            </label>
                        </div>
                    </div>

                    <!-- Row 6: Submit Button -->
                    <div class="portal-col-12">
                        <button type="submit" class="portal-btn portal-btn-success portal-btn-block">
                            Submit Registration &amp; Proceed to Verification
                        </button>
                    </div>
                </div>
            </form>
        </section>
    </div>

    <aside class="portal-layout-sidebar" aria-label="Intake Information and Guidelines">
        <div class="portal-sidebar-card">
            <div class="portal-sidebar-title">
                <span>Public Intake Steps</span>
            </div>
            <ul class="portal-sidebar-list">
                <li>
                    <span class="portal-sidebar-badge">Step 1</span>
                    <span><strong>Submit Details:</strong> Fill out representative details, category, and document subject matter.</span>
                </li>
                <li>
                    <span class="portal-sidebar-badge">Step 2</span>
                    <span><strong>Verify Email:</strong> Enter the 6-digit one-time code sent to your official email.</span>
                </li>
                <li>
                    <span class="portal-sidebar-badge">Step 3</span>
                    <span><strong>Track Online:</strong> Receive your official External Control Number and monitor live progress online.</span>
                </li>
                <li>
                    <span class="portal-sidebar-badge">Step 4</span>
                    <span><strong>Digital Document Release:</strong> Receive your officially approved document and electronic endorsement digitally once processed.</span>
                </li>
            </ul>
        </div>

        <div class="portal-sidebar-card">
            <div class="portal-sidebar-title">
                <span>Office of the Secretary-General</span>
            </div>
            <div class="portal-info-item">
                <span class="portal-info-label">Office:</span>
                Records Section &bull; BTA Parliament Building<br>
                Bangsamoro Government Center, Cotabato City
            </div>
            <div class="portal-info-item">
                <span class="portal-info-label">Operating Hours:</span>
                Monday to Friday: 8:00 AM to 5:00 PM PST<br>
                (Excluding official holidays)
            </div>
        </div>

        <div class="portal-sidebar-card">
            <div class="portal-sidebar-title">
                <span>Submission Requirements</span>
            </div>
            <ul class="portal-sidebar-list">
                <li>
                    <span class="portal-side-mark portal-side-mark-success">&#10003;</span>
                    <span>Valid official email address for authorization</span>
                </li>
                <li>
                    <span class="portal-side-mark portal-side-mark-success">&#10003;</span>
                    <span>Authorized representative contact information</span>
                </li>
                <li>
                    <span class="portal-side-mark portal-side-mark-success">&#10003;</span>
                    <span>Descriptive document subject and classification</span>
                </li>
            </ul>
        </div>

        <div class="portal-sidebar-card">
            <div class="portal-sidebar-title">
                <span>Assistance Desk</span>
            </div>
            <p class="portal-sidebar-text">
                For inquiries regarding tracking, routing verification, or document status:
            </p>
            <div class="portal-info-item">
                <span class="portal-info-label">Email:</span>
                osg.records@bta.gov.ph
            </div>
            <div class="portal-info-item">
                <span class="portal-info-label">Trunkline:</span>
                (064) 552-1234 loc. 104
            </div>
        </div>
    </aside>
</div>
</main>

    <footer class="portal-footer">
        <div class="portal-footer-container">
            <div class="portal-footer-office">
                <strong>Records Section &bull; Office of the Secretary-General</strong><br>
                Bangsamoro Transition Authority Parliament Building, Bangsamoro Government Center, Cotabato City<br>
                Office Hours: Monday to Friday, 8:00 AM to 5:00 PM PST
            </div>
            <div class="portal-footer-legal">
                Official Public Document Portal &bull; Processed in compliance with Republic Act No. 10173 (Data Privacy Act of 2012)
            </div>
        </div>
    </footer>

    <script>
    document.addEventListener('DOMContentLoaded', function() {
        var phoneInput = document.getElementById('phone');
        if (phoneInput) {
            function formatPhone(val) {
                var raw = val.trim();
                var hasPlus = raw.startsWith('+');
                var digits = raw.replace(/\D/g, '');

                if (hasPlus || digits.startsWith('63')) {
                    var sub = digits.startsWith('63') ? digits.slice(2) : digits;
                    var res = '+63';
                    if (sub.length > 0) res += ' ' + sub.slice(0, 3);
                    if (sub.length > 3) res += ' ' + sub.slice(3, 6);
                    if (sub.length > 6) res += ' ' + sub.slice(6, 10);
                    return res;
                } else if (digits.startsWith('09')) {
                    var res = digits.slice(0, 4);
                    if (digits.length > 4) res += ' ' + digits.slice(4, 7);
                    if (digits.length > 7) res += ' ' + digits.slice(7, 11);
                    return res;
                } else if (digits.startsWith('9')) {
                    digits = '0' + digits;
                    var res = digits.slice(0, 4);
                    if (digits.length > 4) res += ' ' + digits.slice(4, 7);
                    if (digits.length > 7) res += ' ' + digits.slice(7, 11);
                    return res;
                } else if (digits.length > 0) {
                    return digits.slice(0, 11);
                }
                return raw;
            }

            phoneInput.addEventListener('input', function() {
                var formatted = formatPhone(this.value);
                if (formatted !== this.value) {
                    this.value = formatted;
                }
            });

            if (phoneInput.value) {
                phoneInput.value = formatPhone(phoneInput.value);
            }
        }
    });
    </script>
    <script src="assets/header-preview.js"></script>
</body>
</html>
