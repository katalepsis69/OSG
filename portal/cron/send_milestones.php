<?php
// BTA OSG External Intake Portal: Hourly Milestone Notification Worker
// Dispatches batch email notifications for newly transitioned document statuses.
// Invoked via system crontab: 0 * * * * php /var/www/portal/cron/send_milestones.php
// Zero third-party dependencies. Uses native cURL to Brevo HTTP API (Port 443).

declare(strict_types=1);

if (php_sapi_name() !== 'cli') {
    http_response_code(403);
    echo "Forbidden: CLI invocation only.\n";
    exit(1);
}

require_once __DIR__ . '/../lib/db.php';
require_once __DIR__ . '/../lib/mail.php';

$startTime = microtime(true);
$logPrefix = '[' . date('Y-m-d H:i:s') . ' UTC] [MilestoneCron] ';

echo $logPrefix . "Starting milestone notification run...\n";

try {
    $pdo = get_db();

    // 1. Fetch unnotified milestones in batches of 50 where attempts < 5
    $stmt = $pdo->prepare(
        'SELECT 
            pm.id AS milestone_id,
            pm.public_status,
            pm.created_at AS milestone_time,
            pm.notification_attempts,
            d.control_number,
            d.document_title,
            r.full_name,
            r.email
         FROM public_milestones pm
         JOIN documents d ON pm.document_id = d.id
         JOIN requesters r ON d.requester_id = r.id
         WHERE pm.notified_at IS NULL AND (pm.notification_attempts IS NULL OR pm.notification_attempts < 5)
         ORDER BY pm.id ASC
         LIMIT 50'
    );
    $stmt->execute();
    $unnotified = $stmt->fetchAll();

    $sentCount = 0;
    $failCount = 0;

    $trackBase = defined('PORTAL_BASE_URL') 
        ? rtrim(PORTAL_BASE_URL, '/') 
        : 'https://portal.bta-osg.gov.ph';

    $updateSuccess = $pdo->prepare('UPDATE public_milestones SET notified_at = NOW(), notification_attempts = notification_attempts + 1 WHERE id = ?');
    $updateFailure = $pdo->prepare('UPDATE public_milestones SET notification_attempts = notification_attempts + 1, last_error = ? WHERE id = ?');

    foreach ($unnotified as $item) {
        $controlNumber = (string)$item['control_number'];
        $status = (string)$item['public_status'];
        $name = (string)$item['full_name'];
        $email = (string)$item['email'];
        $title = (string)$item['document_title'];
        $time = (string)$item['milestone_time'];
        $milestoneId = (int)$item['milestone_id'];

        $trackUrl = $trackBase . '/track.php?cn=' . urlencode($controlNumber);

        $subject = 'Status Update: ' . $controlNumber . ' is now ' . $status . ' - BTA OSG';
        $html = '
        <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; color: #1e293b; line-height: 1.6;">
            <div style="background-color: #0f2a4a; padding: 20px; color: #ffffff; text-align: center;">
                <h2 style="margin: 0; font-size: 1.25rem;">Bangsamoro Transition Authority</h2>
                <p style="margin: 4px 0 0 0; font-size: 0.875rem; color: #cbd5e1;">Office of the Secretary-General</p>
            </div>
            <div style="padding: 24px; border: 1px solid #cbd5e1; background-color: #ffffff;">
                <h3 style="color: #0f2a4a; margin-top: 0;">Document Status Update</h3>
                <p>Dear <strong>' . htmlspecialchars($name, ENT_QUOTES, 'UTF-8') . '</strong>,</p>
                <p>There has been a status update on your submitted document with Control Number <strong>' . htmlspecialchars($controlNumber, ENT_QUOTES, 'UTF-8') . '</strong>.</p>
                <div style="margin: 20px 0; padding: 16px; background-color: #f1f5f9; border-left: 4px solid #0f2a4a; border-radius: 4px;">
                    <p style="margin: 0 0 8px 0; font-size: 0.875rem; color: #64748b;">CURRENT PUBLIC STATUS</p>
                    <p style="margin: 0; font-size: 1.25rem; font-weight: bold; color: #0f2a4a;">' . htmlspecialchars($status, ENT_QUOTES, 'UTF-8') . '</p>
                </div>
                <table style="width: 100%; border-collapse: collapse; margin: 20px 0; font-size: 0.9rem;">
                    <tr>
                        <td style="padding: 8px; border-bottom: 1px solid #e2e8f0; color: #64748b; width: 140px;">Title / Subject:</td>
                        <td style="padding: 8px; border-bottom: 1px solid #e2e8f0; font-weight: bold;">' . htmlspecialchars($title, ENT_QUOTES, 'UTF-8') . '</td>
                    </tr>
                    <tr>
                        <td style="padding: 8px; border-bottom: 1px solid #e2e8f0; color: #64748b;">Update Recorded:</td>
                        <td style="padding: 8px; border-bottom: 1px solid #e2e8f0;">' . htmlspecialchars($time, ENT_QUOTES, 'UTF-8') . ' UTC</td>
                    </tr>
                </table>
                <p style="margin: 24px 0; text-align: center;">
                    <a href="' . htmlspecialchars($trackUrl, ENT_QUOTES, 'UTF-8') . '" style="display: inline-block; background-color: #0f2a4a; color: #ffffff; padding: 10px 20px; text-decoration: none; border-radius: 4px; font-weight: bold;">
                        View Full Milestone Timeline
                    </a>
                </p>
                <hr style="border: none; border-top: 1px solid #e2e8f0; margin: 20px 0;">
                <p style="font-size: 0.8125rem; color: #94a3b8;">This is an automated notification from the BTA OSG Public Document Portal.</p>
            </div>
        </div>';

        $res = send_brevo_mail($email, $name, $subject, $html);
        if ($res['success']) {
            $updateSuccess->execute([$milestoneId]);
            $sentCount++;
        } else {
            $failCount++;
            $errMsg = substr((string)($res['error'] ?? 'Unknown send error'), 0, 255);
            $updateFailure->execute([$errMsg, $milestoneId]);
            // The cron log is world-readable on many VPS images: recipient addresses are
            // PII, so failures identify the milestone, not the citizen.
            echo $logPrefix . "Failed milestone #$milestoneId for $controlNumber: " . $errMsg . "\n";
        }
    }

    // 2. Housekeeping: Expire unverified submissions older than 24 hours
    $pruneSub = $pdo->prepare(
        'UPDATE pending_submissions
         SET status = \'expired\'
         WHERE status = \'pending\' AND created_at < DATE_SUB(NOW(), INTERVAL 24 HOUR)'
    );
    $pruneSub->execute();
    $expiredCount = $pruneSub->rowCount();

    // 3. Housekeeping: Remove expired OTP codes older than 7 days, plus any dev-helper
    // OTP temp files for tokens that no longer have a pending submission.
    $pruneOtp = $pdo->prepare('DELETE FROM otp_codes WHERE created_at < DATE_SUB(NOW(), INTERVAL 7 DAY)');
    $pruneOtp->execute();
    $prunedOtpCount = $pruneOtp->rowCount();

    $otpTmpDir = sys_get_temp_dir();
    foreach (glob($otpTmpDir . DIRECTORY_SEPARATOR . 'bta_dev_otp_*') ?: [] as $otpFile) {
        $otpToken = substr(basename((string)$otpFile), strlen('bta_dev_otp_'));
        if (!preg_match('/^[a-f0-9]{32}$/', $otpToken)) {
            @unlink($otpFile);
            continue;
        }
        $liveStmt = $pdo->prepare("SELECT COUNT(*) FROM pending_submissions WHERE id = ? AND status = 'pending'");
        $liveStmt->execute([$otpToken]);
        if ((int)$liveStmt->fetchColumn() === 0) {
            @unlink($otpFile);
        }
    }

    $elapsed = round(microtime(true) - $startTime, 3);
    echo $logPrefix . sprintf(
        "Run complete in %s s. Sent: %d, Failed: %d, Expired submissions: %d, Pruned OTPs: %d\n",
        $elapsed,
        $sentCount,
        $failCount,
        $expiredCount,
        $prunedOtpCount
    );

} catch (Throwable $ex) {
    echo $logPrefix . 'Fatal error during milestone notification run: ' . $ex->getMessage() . "\n";
    exit(1);
}
