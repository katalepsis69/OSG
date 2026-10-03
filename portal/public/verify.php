<?php
// BTA OSG External Intake Portal: Email OTP Verification & Receipt
// Bangsamoro Transition Authority - Office of the Secretary-General
// Verifies 6-digit code, assigns unique monotonic control number, sends receipt.

declare(strict_types=1);

require_once __DIR__ . '/../lib/db.php';
require_once __DIR__ . '/../lib/mail.php';
require_once __DIR__ . '/../lib/ratelimit.php';
require_once __DIR__ . '/../lib/session.php';
portal_session_start();

$token = trim((string)($_GET['token'] ?? $_POST['token'] ?? $_SESSION['pending_token'] ?? ''));
if ($token === '' || !preg_match('/^[a-f0-9]{32}$/', $token)) {
    header('Location: index.php?notice=expired');
    exit;
}

$pdo = get_db();

// 1. Fetch pending submission
$subStmt = $pdo->prepare('SELECT * FROM pending_submissions WHERE id = ?');
$subStmt->execute([$token]);
$submission = $subStmt->fetch();

if (!$submission) {
    header('Location: index.php?notice=expired');
    exit;
}

$errorMessage = null;
$successMessage = null;
$verifiedDocument = null;
$receiptMailOk = null; // null = no send attempt this request (revisit); true/false = actual result

// Suffix alphabet: 31 characters (excluding I, L, O, 0, 1 for visual readability)
const SUFFIX_ALPHABET = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789';

/**
 * Generates a 4-character random suffix from the 31-character alphabet.
 */
function generate_suffix(): string {
    $alpha = SUFFIX_ALPHABET;
    $len = strlen($alpha);
    $res = '';
    for ($i = 0; $i < 4; $i++) {
        $res .= $alpha[random_int(0, $len - 1)];
    }
    return $res;
}

// 2. Handle Resend OTP Request
if ($_SERVER['REQUEST_METHOD'] === 'POST' && (isset($_POST['action']) && $_POST['action'] === 'resend')) {
    portal_csrf_verify($_POST['csrf_token'] ?? null);
    try {
        if (!osg_rate_limit('resend', 10, 3600)) {
            $errorMessage = 'Too many code requests from your IP address. Please try again later.';
        } elseif ($submission['status'] !== 'pending') {
            $errorMessage = 'This submission has already been processed.';
        } else {
            // Check rate limiting / cooldown (60 seconds)
            $checkStmt = $pdo->prepare(
                'SELECT created_at FROM otp_codes WHERE submission_id = ? ORDER BY id DESC LIMIT 1'
            );
            $checkStmt->execute([$token]);
            $lastOtpTime = $checkStmt->fetchColumn();

            $canResend = true;
            if ($lastOtpTime) {
                $secondsSince = time() - strtotime((string)$lastOtpTime);
                if ($secondsSince < 60) {
                    $canResend = false;
                    $errorMessage = 'Please wait ' . (60 - $secondsSince) . ' seconds before requesting another code.';
                }
            }

            if ($canResend) {
                $newOtp = (string)random_int(100000, 999999);
                $newHash = password_hash($newOtp, PASSWORD_DEFAULT);

                $subject = 'Your new BTA OSG verification code';
                $html = '
                <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; color: #1e293b; line-height: 1.6;">
                    <div style="background-color: #0f2a4a; padding: 20px; color: #ffffff; text-align: center;">
                        <h2 style="margin: 0; font-size: 1.25rem;">Bangsamoro Transition Authority</h2>
                        <p style="margin: 4px 0 0 0; font-size: 0.875rem; color: #cbd5e1;">Office of the Secretary-General</p>
                    </div>
                    <div style="padding: 24px; border: 1px solid #cbd5e1; background-color: #ffffff;">
                        <p>Dear <strong>' . htmlspecialchars($submission['full_name'], ENT_QUOTES, 'UTF-8') . '</strong>,</p>
                        <p>Your requested replacement verification code is below:</p>
                        <div style="margin: 24px 0; text-align: center;">
                            <span style="display: inline-block; font-family: monospace; font-size: 2rem; font-weight: bold; letter-spacing: 0.5rem; background-color: #f1f5f9; padding: 12px 24px; border: 1px solid #cbd5e1; border-radius: 6px; color: #0f2a4a;">
                                ' . htmlspecialchars($newOtp, ENT_QUOTES, 'UTF-8') . '
                            </span>
                        </div>
                        <p style="font-size: 0.875rem; color: #64748b;">Valid for 10 minutes. Please enter this code on the verification screen.</p>
                    </div>
                </div>';

                $mailRes = send_brevo_mail($submission['email'], $submission['full_name'], $subject, $html);
                if ($mailRes['success']) {
                    // Stored only once the new code is actually delivered. Verifying reads the
                    // newest unconsumed row, so inserting before a failed send would discard the
                    // code the user already has without replacing it.
                    $insStmt = $pdo->prepare(
                        'INSERT INTO otp_codes (submission_id, code_hash, expires_at, attempts, created_at)
                         VALUES (?, ?, DATE_ADD(NOW(), INTERVAL 10 MINUTE), 0, NOW())'
                    );
                    $insStmt->execute([$token, $newHash]);
                    // Dev-only helper storage (see ALLOW_DEV_OTP_DISPLAY); off in production.
                    if (defined('ALLOW_DEV_OTP_DISPLAY') && ALLOW_DEV_OTP_DISPLAY === true) {
                        $_SESSION['dev_last_otp'] = $newOtp;
                        $devOtpPath = sys_get_temp_dir() . DIRECTORY_SEPARATOR . 'bta_dev_otp_' . $token;
                        @file_put_contents($devOtpPath, $newOtp);
                        @chmod($devOtpPath, 0600);
                    }

                    $successMessage = 'A new 6-digit verification code has been dispatched to your email.';
                } else {
                    error_log('[Mail Warning] Resend OTP dispatch failed: ' . ($mailRes['error'] ?? 'unknown'));
                    $errorMessage = 'We could not deliver a new code to your email just now. Please try again, or contact the Records Section.';
                }
            }
        }
    } catch (Throwable $ex) {
        error_log('[OTP Resend Error] ' . $ex->getMessage());
        $errorMessage = 'Failed to generate a new verification code. Please try again.';
    }
}

// 3. Handle Verify Code Request
if ($_SERVER['REQUEST_METHOD'] === 'POST' && (!isset($_POST['action']) || $_POST['action'] === 'verify')) {
    portal_csrf_verify($_POST['csrf_token'] ?? null);
    $enteredCode = trim((string)($_POST['otp_code'] ?? ''));

    if (!preg_match('/^\d{6}$/', $enteredCode)) {
        $errorMessage = 'Please enter a valid 6-digit numeric verification code.';
    } else {
        $otpAlreadyUsed = false;
        try {
            // Find active unconsumed OTP
            $otpStmt = $pdo->prepare(
                'SELECT * FROM otp_codes 
                 WHERE submission_id = ? AND consumed_at IS NULL 
                 ORDER BY id DESC LIMIT 1'
            );
            $otpStmt->execute([$token]);
            $otpRecord = $otpStmt->fetch();

            if (!$otpRecord) {
                $errorMessage = 'No active verification code found. Please request a new code.';
            } elseif ((int)$otpRecord['attempts'] >= 5) {
                $errorMessage = 'Maximum verification attempts exceeded. Please request a new code.';
            } elseif (strtotime((string)$otpRecord['expires_at']) < time()) {
                // Expired: drop any dev-helper copy of the dead code so it cannot linger.
                @unlink(sys_get_temp_dir() . DIRECTORY_SEPARATOR . 'bta_dev_otp_' . $token);
                $errorMessage = 'This verification code has expired. Please request a new code.';
            } elseif (!password_verify($enteredCode, (string)$otpRecord['code_hash'])) {
                // Increment failed attempt count
                $pdo->prepare('UPDATE otp_codes SET attempts = attempts + 1 WHERE id = ?')->execute([$otpRecord['id']]);
                $remaining = 4 - (int)$otpRecord['attempts'];
                if ($remaining > 0) {
                    $errorMessage = 'Incorrect verification code. ' . $remaining . ' attempt(s) remaining.';
                } else {
                    $errorMessage = 'Incorrect code. Maximum verification attempts reached. Please request a new code.';
                }
            } else {
                // OTP verified successfully: Execute atomic registration transaction
                unset($_SESSION['dev_last_otp']);
                @unlink(sys_get_temp_dir() . DIRECTORY_SEPARATOR . 'bta_dev_otp_' . $token);
                $pdo->beginTransaction();

                // 1. Mark OTP as consumed. The conditional guard makes a double submit
                // from two concurrent correct entries single-winner: the loser's update
                // matches no row and the transaction aborts before any counter moves.
                $consumeStmt = $pdo->prepare('UPDATE otp_codes SET consumed_at = NOW() WHERE id = ? AND consumed_at IS NULL');
                $consumeStmt->execute([$otpRecord['id']]);
                if ($consumeStmt->rowCount() !== 1) {
                    $otpAlreadyUsed = true;
                    throw new RuntimeException('OTP was consumed by a concurrent verification.');
                }

                // 2. Fetch and lock counter sequence
                $category = (string)$submission['category'];
                $year = (int)date('Y');

                $counterStmt = $pdo->prepare(
                    'SELECT seq FROM counters WHERE prefix = ? AND yr = ? FOR UPDATE'
                );
                $counterStmt->execute([$category, $year]);
                $currentSeq = $counterStmt->fetchColumn();

                if ($currentSeq === false) {
                    // Initialize counter for prefix and year if missing
                    $pdo->prepare('INSERT INTO counters (prefix, yr, seq) VALUES (?, ?, 0)')->execute([$category, $year]);
                    $currentSeq = 0;
                }

                $nextSeq = (int)$currentSeq + 1;
                $pdo->prepare('UPDATE counters SET seq = ? WHERE prefix = ? AND yr = ?')->execute([$nextSeq, $category, $year]);

                // 3. Generate Control Number: {PREFIX}-{YYYY}-{SEQ:04d}-{SUFFIX}
                $suffix = generate_suffix();
                $controlNumber = sprintf('%s-%d-%04d-%s', $category, $year, $nextSeq, $suffix);

                // 4. Upsert Requester
                $reqStmt = $pdo->prepare('SELECT id FROM requesters WHERE email = ?');
                $reqStmt->execute([$submission['email']]);
                $requesterId = $reqStmt->fetchColumn();
                $gender = (string)($submission['gender'] ?? 'Prefer not to say');

                if ($requesterId === false) {
                    $insReq = $pdo->prepare(
                        'INSERT INTO requesters (full_name, email, phone, gender, created_at) VALUES (?, ?, ?, ?, NOW())'
                    );
                    $insReq->execute([$submission['full_name'], $submission['email'], $submission['phone'], $gender]);
                    $requesterId = (int)$pdo->lastInsertId();
                } else {
                    $requesterId = (int)$requesterId;
                    // Keep contact details and gender demographic fresh
                    $updReq = $pdo->prepare('UPDATE requesters SET full_name = ?, phone = ?, gender = ? WHERE id = ?');
                    $updReq->execute([$submission['full_name'], $submission['phone'], $gender, $requesterId]);
                }

                // 5. Insert Document
                $docStmt = $pdo->prepare(
                    'INSERT INTO documents (control_number, category, document_title, requester_id, public_status, created_at)
                     VALUES (?, ?, ?, ?, "Received", NOW())'
                );
                $docStmt->execute([$controlNumber, $category, $submission['document_title'], $requesterId]);
                $documentId = (int)$pdo->lastInsertId();

                // 6. Insert initial Received milestone
                $msStmt = $pdo->prepare(
                    'INSERT INTO public_milestones (document_id, public_status, created_at, notified_at)
                     VALUES (?, "Received", NOW(), NOW())'
                );
                $msStmt->execute([$documentId]);

                // 7. Mark pending submission verified
                $pdo->prepare('UPDATE pending_submissions SET status = "verified" WHERE id = ?')->execute([$token]);

                $pdo->commit();

                // 8. Dispatch Confirmation Email with Control Number and Tracking Link
                $trackUrl = defined('PORTAL_BASE_URL') 
                    ? rtrim(PORTAL_BASE_URL, '/') . '/track.php?cn=' . urlencode($controlNumber)
                    : 'https://portal.bta-osg.gov.ph/track.php?cn=' . urlencode($controlNumber);

                $confSubject = 'Official Receipt: ' . $controlNumber . ' - BTA OSG Document Submission';
                $confHtml = '
                <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; color: #1e293b; line-height: 1.6;">
                    <div style="background-color: #0f2a4a; padding: 24px; color: #ffffff; text-align: center;">
                        <h2 style="margin: 0; font-size: 1.25rem;">Bangsamoro Transition Authority</h2>
                        <p style="margin: 4px 0 0 0; font-size: 0.875rem; color: #cbd5e1;">Office of the Secretary-General</p>
                    </div>
                    <div style="padding: 24px; border: 1px solid #cbd5e1; background-color: #ffffff;">
                        <h3 style="color: #0f2a4a; margin-top: 0;">Official Submission Receipt</h3>
                        <p>Dear <strong>' . htmlspecialchars($submission['full_name'], ENT_QUOTES, 'UTF-8') . '</strong>,</p>
                        <p>Your document submission has been verified and registered into the BTA OSG Intake Registry. Your official external control number is:</p>
                        <div style="margin: 20px 0; text-align: center;">
                            <span style="display: inline-block; font-family: monospace; font-size: 1.5rem; font-weight: bold; background-color: #f1f5f9; padding: 12px 20px; border: 2px solid #0f2a4a; border-radius: 6px; color: #0f2a4a;">
                                ' . htmlspecialchars($controlNumber, ENT_QUOTES, 'UTF-8') . '
                            </span>
                        </div>
                        <table style="width: 100%; border-collapse: collapse; margin: 20px 0; font-size: 0.9rem;">
                            <tr>
                                <td style="padding: 8px; border-bottom: 1px solid #e2e8f0; color: #64748b; width: 140px;">Title / Subject:</td>
                                <td style="padding: 8px; border-bottom: 1px solid #e2e8f0; font-weight: bold;">' . htmlspecialchars($submission['document_title'], ENT_QUOTES, 'UTF-8') . '</td>
                            </tr>
                            <tr>
                                <td style="padding: 8px; border-bottom: 1px solid #e2e8f0; color: #64748b;">Category:</td>
                                <td style="padding: 8px; border-bottom: 1px solid #e2e8f0;">' . htmlspecialchars($category, ENT_QUOTES, 'UTF-8') . '</td>
                            </tr>
                            <tr>
                                <td style="padding: 8px; border-bottom: 1px solid #e2e8f0; color: #64748b;">Date Registered:</td>
                                <td style="padding: 8px; border-bottom: 1px solid #e2e8f0;">' . date('Y-m-d H:i:s') . ' UTC</td>
                            </tr>
                            <tr>
                                <td style="padding: 8px; border-bottom: 1px solid #e2e8f0; color: #64748b;">Initial Status:</td>
                                <td style="padding: 8px; border-bottom: 1px solid #e2e8f0; color: #0369a1; font-weight: bold;">Received</td>
                            </tr>
                        </table>
                        <p>You can check the live tracking status of your document at any time using the link below:</p>
                        <p style="margin: 20px 0; text-align: center;">
                            <a href="' . htmlspecialchars($trackUrl, ENT_QUOTES, 'UTF-8') . '" style="display: inline-block; background-color: #0f2a4a; color: #ffffff; padding: 10px 20px; text-decoration: none; border-radius: 4px; font-weight: bold;">
                                Track Document Online
                            </a>
                        </p>
                        <p style="font-size: 0.8125rem; color: #64748b;">Please preserve your Control Number for all future inquiries with the BTA OSG Records Section.</p>
                    </div>
                </div>';

                $receiptMail = send_brevo_mail($submission['email'], $submission['full_name'], $confSubject, $confHtml);
                $receiptMailOk = $receiptMail['success'];
                if (!$receiptMailOk) {
                    error_log('[Mail Warning] Receipt dispatch failed: ' . ($receiptMail['error'] ?? 'unknown'));
                }

                // Set verified document payload for receipt display
                $verifiedDocument = [
                    'control_number' => $controlNumber,
                    'document_title' => $submission['document_title'],
                    'category' => $category,
                    'full_name' => $submission['full_name'],
                    'email' => $submission['email'],
                    'phone' => $submission['phone'],
                    'created_at' => date('Y-m-d H:i:s') . ' UTC',
                    'track_url' => $trackUrl
                ];
            }
        } catch (Throwable $ex) {
            if ($pdo && $pdo->inTransaction()) {
                $pdo->rollBack();
            }
            error_log('[Verification Error] ' . $ex->getMessage());
            $errorMessage = 'A server error occurred during verification. Please try again.';
        }
        if ($otpAlreadyUsed) {
            $errorMessage = 'This verification code was just used. Please request a new code.';
        }
    }
}

// 4. If submission was already verified in a prior request, load existing document
if (!$verifiedDocument && $submission['status'] === 'verified') {
    $existingStmt = $pdo->prepare(
        'SELECT d.*, r.full_name, r.email, r.phone 
         FROM documents d 
         JOIN requesters r ON d.requester_id = r.id 
         WHERE r.email = ? AND d.document_title = ? 
         ORDER BY d.id DESC LIMIT 1'
    );
    $existingStmt->execute([$submission['email'], $submission['document_title']]);
    $existingDoc = $existingStmt->fetch();
    if ($existingDoc) {
        $trackUrl = defined('PORTAL_BASE_URL')
            ? rtrim(PORTAL_BASE_URL, '/') . '/track.php?cn=' . urlencode((string)$existingDoc['control_number'])
            : 'https://portal.bta-osg.gov.ph/track.php?cn=' . urlencode((string)$existingDoc['control_number']);

        $verifiedDocument = [
            'control_number' => $existingDoc['control_number'],
            'document_title' => $existingDoc['document_title'],
            'category' => $existingDoc['category'],
            'full_name' => $existingDoc['full_name'],
            'email' => $existingDoc['email'],
            'phone' => $existingDoc['phone'],
            'created_at' => $existingDoc['created_at'] . ' UTC',
            'track_url' => $trackUrl
        ];
    }
}
$activePage = 'index';
// The on-page OTP display is gated ONLY by the explicit config flag
// ALLOW_DEV_OTP_DISPLAY (default false). It must never key on REMOTE_ADDR or the base
// URL: behind the Cloudflare tunnel every request arrives with REMOTE_ADDR=127.0.0.1,
// so a localhost check would print the live verification code to the whole internet.
$devOtp = null;
if (defined('ALLOW_DEV_OTP_DISPLAY') && ALLOW_DEV_OTP_DISPLAY === true) {
    $devOtp = $_SESSION['dev_last_otp'] ?? null;
    if (!$devOtp) {
        $devOtpFile = sys_get_temp_dir() . DIRECTORY_SEPARATOR . 'bta_dev_otp_' . $token;
        if (file_exists($devOtpFile)) {
            $devOtp = trim((string)@file_get_contents($devOtpFile));
        }
    }
}

/**
 * Formats a UTC timestamp into Philippine Standard Time (PST / UTC+8)
 */
function format_pst_time(string $utcTimeStr): string {
    try {
        $dt = new DateTime($utcTimeStr, new DateTimeZone('UTC'));
        $dt->setTimezone(new DateTimeZone('Asia/Manila'));
        return $dt->format('F j, Y \a\t g:i A \P\S\T');
    } catch (Throwable $e) {
        return htmlspecialchars($utcTimeStr, ENT_QUOTES, 'UTF-8') . ' UTC';
    }
}
?>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Verify Submission - OSGPortal Public Intake</title>
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
                <?php if ($verifiedDocument): ?>
                    <!-- Official Receipt Card -->
                    <section class="portal-receipt">
                        <div class="portal-stepper" aria-label="Intake Journey Progress">
                            <div class="portal-step-item completed">
                                <span class="portal-step-num">&#10003;</span>
                                <span>Submit Details</span>
                            </div>
                            <div class="portal-step-item completed">
                                <span class="portal-step-num">&#10003;</span>
                                <span>Verify Email (OTP)</span>
                            </div>
                            <div class="portal-step-item completed">
                                <span class="portal-step-num">&#10003;</span>
                                <span>Track Online &amp; Receipt</span>
                            </div>
                            <div class="portal-step-item active">
                                <span class="portal-step-num">4</span>
                                <span>Digital Document Release</span>
                            </div>
                        </div>

                        <div style="border-bottom: 2px solid var(--color-primary); padding-bottom: 1rem; margin-bottom: 1.5rem;">
                            <span style="font-size: 0.8125rem; text-transform: uppercase; letter-spacing: 0.05em; color: var(--color-accent); font-weight: 700;">
                                Bangsamoro Transition Authority &bull; Office of the Secretary-General
                            </span>
                            <h1 style="color: var(--color-primary); font-size: 1.5rem; margin-top: 0.25rem;">
                                Official Submission Receipt
                            </h1>
                        </div>

                <div class="portal-alert portal-alert-success" role="alert">
                    <span class="portal-alert-title">Submission Successfully Verified</span>
                    Your document has been registered into the public digital intake queue.
                    <?php if ($receiptMailOk === null): ?>
                        This is a copy of your receipt. Your submission was already verified earlier.
                    <?php elseif ($receiptMailOk): ?>
                        A confirmation copy has been sent to your email.
                    <?php else: ?>
                        We could not email you a copy, so please save or print this receipt now. Your control number remains valid.
                    <?php endif; ?>
                </div>

                <div style="text-align: center; margin: 1.5rem 0;">
                    <div style="font-size: 0.875rem; color: var(--color-text-muted); font-weight: 600;">OFFICIAL CONTROL NUMBER</div>
                    <div class="portal-control-number-badge">
                        <?= htmlspecialchars((string)$verifiedDocument['control_number'], ENT_QUOTES, 'UTF-8') ?>
                    </div>
                    <div>
                        <span class="portal-status-pill portal-status-received">Status: Received</span>
                    </div>
                </div>

                <div class="portal-kv-list">
                    <div class="portal-kv-label">Document Title:</div>
                    <div class="portal-kv-value" style="font-weight: 600;">
                        <?= htmlspecialchars((string)$verifiedDocument['document_title'], ENT_QUOTES, 'UTF-8') ?>
                    </div>

                    <div class="portal-kv-label">Category:</div>
                    <div class="portal-kv-value">
                        <?= htmlspecialchars((string)$verifiedDocument['category'], ENT_QUOTES, 'UTF-8') ?>
                    </div>

                    <div class="portal-kv-label">Submitter:</div>
                    <div class="portal-kv-value">
                        <?= htmlspecialchars((string)$verifiedDocument['full_name'], ENT_QUOTES, 'UTF-8') ?>
                    </div>

                    <div class="portal-kv-label">Email:</div>
                    <div class="portal-kv-value">
                        <?= htmlspecialchars((string)$verifiedDocument['email'], ENT_QUOTES, 'UTF-8') ?>
                    </div>

                    <div class="portal-kv-label">Phone:</div>
                    <div class="portal-kv-value">
                        <?= htmlspecialchars((string)$verifiedDocument['phone'], ENT_QUOTES, 'UTF-8') ?>
                    </div>

                    <div class="portal-kv-label">Registered Date:</div>
                    <div class="portal-kv-value">
                        <?= format_pst_time((string)$verifiedDocument['created_at']) ?>
                    </div>
                </div>

                <div class="portal-actions-bar">
                    <a href="<?= htmlspecialchars((string)$verifiedDocument['track_url'], ENT_QUOTES, 'UTF-8') ?>" class="portal-btn portal-btn-primary">
                        Track Document Status
                    </a>
                    <button type="button" class="portal-btn portal-btn-secondary" onclick="window.print();">
                        Print Official Receipt
                    </button>
                    <a href="index.php" class="portal-btn portal-btn-secondary">
                        Submit Another Document
                    </a>
                </div>
            </section>
        <?php else: ?>
            <!-- 6-Digit OTP Verification Form -->
            <section class="portal-card">
                <div class="portal-stepper" aria-label="Intake Journey Progress">
                    <div class="portal-step-item completed">
                        <span class="portal-step-num">&#10003;</span>
                        <span>Submit Details</span>
                    </div>
                    <div class="portal-step-item active">
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
                    <h1 class="portal-card-title">Verify Email Address</h1>
                    <p class="portal-card-subtitle">
                        A 6-digit verification code was sent to <strong><?= htmlspecialchars((string)$submission['email'], ENT_QUOTES, 'UTF-8') ?></strong>.
                    </p>
                </div>

                <div class="portal-alert portal-alert-info" role="status">
                    <span class="portal-alert-title">Email Delivery Notice:</span>
                    Your 6-digit code has been dispatched via Brevo to <strong><?= htmlspecialchars((string)$submission['email'], ENT_QUOTES, 'UTF-8') ?></strong>. If it is not in your Inbox within a minute, <strong>please check your Spam or Junk folder</strong> (email providers such as Gmail and Yahoo frequently route automated relay messages to Spam).
                </div>

                <?php if (!empty($devOtp)): ?>
                    <div class="portal-alert portal-alert-warning" role="status" style="font-family: var(--font-mono); font-size: 0.875rem;">
                        <span class="portal-alert-title">Local Development Helper:</span>
                        Current verification code: <strong style="font-size: 1.15rem; letter-spacing: 0.15em; color: var(--color-primary);"><?= htmlspecialchars((string)$devOtp, ENT_QUOTES, 'UTF-8') ?></strong>
                        <div style="font-family: var(--font-sans); font-size: 0.8125rem; color: var(--color-text-muted); margin-top: 0.25rem;">
                            Visible in local development environment. Email is also dispatched via Brevo.
                        </div>
                    </div>
                <?php endif; ?>

                <?php if ($errorMessage): ?>
                    <div class="portal-alert portal-alert-danger" role="alert">
                        <span class="portal-alert-title">Verification Issue:</span>
                        <?= htmlspecialchars($errorMessage, ENT_QUOTES, 'UTF-8') ?>
                    </div>
                <?php endif; ?>

                <?php if ($successMessage): ?>
                    <div class="portal-alert portal-alert-success" role="alert">
                        <span class="portal-alert-title">Code Dispatched:</span>
                        <?= htmlspecialchars($successMessage, ENT_QUOTES, 'UTF-8') ?>
                    </div>
                <?php endif; ?>

                <form action="verify.php" method="POST"
                      onsubmit="if (this.dataset.busy) { return false; } this.dataset.busy = '1'; var b = this.querySelector('button[type=submit]'); if (b) { b.disabled = true; b.textContent = 'Verifying...'; }">
                    <input type="hidden" name="token" value="<?= htmlspecialchars($token, ENT_QUOTES, 'UTF-8') ?>">
                    <input type="hidden" name="action" value="verify">
                    <input type="hidden" name="csrf_token" value="<?= htmlspecialchars(portal_csrf_token(), ENT_QUOTES, 'UTF-8') ?>">

                    <div class="portal-otp-container">
                        <label for="otp_code" class="portal-label" style="margin-bottom: 0.75rem;">
                            Enter 6-Digit Verification Code
                        </label>
                        <input type="text" id="otp_code" name="otp_code" class="portal-otp-input"
                               inputmode="numeric" pattern="[0-9]{6}" maxlength="6" autocomplete="one-time-code"
                               required autofocus placeholder="------">
                        <p class="portal-help-text" style="margin-top: 0.5rem;">
                            The code expires in 10 minutes.
                        </p>
                    </div>

                    <div style="margin-top: 1.5rem;">
                        <button type="submit" class="portal-btn portal-btn-primary portal-btn-block">
                            Verify &amp; Issue Control Number
                        </button>
                    </div>
                </form>

                <div style="margin-top: 2rem; padding-top: 1.5rem; border-top: 1px solid var(--color-border); text-align: center;">
                    <form action="verify.php" method="POST" id="resendForm"
                          onsubmit="if (this.dataset.busy) { return false; } this.dataset.busy = '1'; var b = this.querySelector('button[type=submit]'); if (b) { b.disabled = true; b.textContent = 'Sending...'; }">
                        <input type="hidden" name="token" value="<?= htmlspecialchars($token, ENT_QUOTES, 'UTF-8') ?>">
                        <input type="hidden" name="action" value="resend">
                        <input type="hidden" name="csrf_token" value="<?= htmlspecialchars(portal_csrf_token(), ENT_QUOTES, 'UTF-8') ?>">
                        <p style="font-size: 0.875rem; color: var(--color-text-muted);">
                            Didn't receive the email code? Check your spam folder or
                        </p>
                        <button type="submit" class="portal-btn portal-btn-secondary" style="margin-top: 0.5rem; font-size: 0.875rem;">
                            Resend Verification Code
                        </button>
                    </form>
                </div>
            </section>
        <?php endif; ?>
    </div>

    <aside class="portal-layout-sidebar" aria-label="Verification and Official Information">
        <?php if ($verifiedDocument): ?>
            <!-- Receipt View Sidebar: Digital Submission Confirmed -->
            <div class="portal-sidebar-card">
                <div class="portal-sidebar-title">
                    <span>Digital Submission Confirmed</span>
                </div>
                <ul class="portal-sidebar-list">
                    <li>
                        <span class="portal-side-mark portal-side-mark-success">&#10003;</span>
                        <span>External Control Number officially assigned</span>
                    </li>
                    <li>
                        <span class="portal-side-mark portal-side-mark-success">&#10003;</span>
                        <span>Document entered in Records Section digital queue</span>
                    </li>
                    <li>
                        <span class="portal-side-mark portal-side-mark-success">&#10003;</span>
                        <span>Official confirmation receipt dispatched to email</span>
                    </li>
                    <li>
                        <span class="portal-side-mark portal-side-mark-success">&#10003;</span>
                        <span>Live status tracking available 24/7 online</span>
                    </li>
                    <li>
                        <span class="portal-side-mark portal-side-mark-success">&#10003;</span>
                        <span>Approved documents released digitally to registered email</span>
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
        <?php else: ?>
            <!-- OTP Entry View Sidebar -->
            <div class="portal-sidebar-card">
                <div class="portal-sidebar-title">
                    <span>Identity Verification</span>
                </div>
                <p class="portal-sidebar-text">
                    For public records integrity and Data Privacy Act (RA 10173) compliance, submissions require one-time email authorization.
                </p>
                <ul class="portal-sidebar-list">
                    <li>
                        <span class="portal-side-mark portal-side-mark-primary">&bull;</span>
                        <span>Verification codes remain valid for 10 minutes.</span>
                    </li>
                    <li>
                        <span class="portal-side-mark portal-side-mark-primary">&bull;</span>
                        <span>Account rate limiting enforces a maximum of 5 attempts.</span>
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
                    <span>Next Steps</span>
                </div>
                <p class="portal-sidebar-text">
                    Upon entering your valid 6-digit code:
                </p>
                <ul class="portal-sidebar-list">
                    <li>
                        <span class="portal-side-mark portal-side-mark-success">1.</span>
                        <span>Your unique External Control Number is generated.</span>
                    </li>
                    <li>
                        <span class="portal-side-mark portal-side-mark-success">2.</span>
                        <span>An official digital receipt is displayed and emailed to you.</span>
                    </li>
                    <li>
                        <span class="portal-side-mark portal-side-mark-success">3.</span>
                        <span>Track your document status online at any time.</span>
                    </li>
                    <li>
                        <span class="portal-side-mark portal-side-mark-success">4.</span>
                        <span>Receive your officially approved document and electronic endorsement digitally.</span>
                    </li>
                </ul>
            </div>

            <div class="portal-sidebar-card">
                <div class="portal-sidebar-title">
                    <span>Assistance Desk</span>
                </div>
                <div class="portal-info-item">
                    <span class="portal-info-label">Email:</span>
                    osg.records@bta.gov.ph
                </div>
                <div class="portal-info-item">
                    <span class="portal-info-label">Trunkline:</span>
                    (064) 552-1234 loc. 104
                </div>
            </div>
        <?php endif; ?>
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
    <script src="assets/header-preview.js"></script>
</body>
</html>
