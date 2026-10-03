<?php
// BTA OSG External Intake Portal: Public Status Tracking
// Bangsamoro Transition Authority - Office of the Secretary-General
// Queries strictly public_documents projection view and public_milestones.
// Zero internal identifiers or confidential routing notes exposed.

declare(strict_types=1);

require_once __DIR__ . '/../lib/db.php';
require_once __DIR__ . '/../lib/ratelimit.php';
require_once __DIR__ . '/../lib/session.php';
portal_session_start();

$searchedCn = trim((string)($_GET['cn'] ?? $_POST['cn'] ?? ''));
$document = null;
$milestones = [];
$searched = false;
$errorMessage = null;

// IP Rate Limiting: maximum 10 queries per minute per client. Bucketing lives in
// lib/ratelimit.php, which honours CF-Connecting-IP only over loopback (tunnel
// traffic) and never trusts it from a direct cloud connection.
if ($searchedCn !== '') {
    $searched = true;

    if (!osg_rate_limit('track', 10, 60)) {
        $errorMessage = 'Too many tracking requests from your IP address. Please wait a minute before searching again.';
    } else {

        // Normalize format: uppercase, trimmed, and auto-hyphenate if hyphens were omitted
        $rawClean = strtoupper(preg_replace('/[^A-Za-z0-9]/', '', $searchedCn));
        if (preg_match('/^(COMM|LEG|FIN|TO)(\d{4})(\d{4})([A-Z0-9]{4})$/', $rawClean, $m)) {
            $normalizedCn = $m[1] . '-' . $m[2] . '-' . $m[3] . '-' . $m[4];
        } else {
            $normalizedCn = strtoupper(trim($searchedCn));
        }

        // Control number pattern: e.g. COMM-2026-0001-K9X2
        if (!preg_match('/^[A-Z]{2,4}-\d{4}-\d{4}-[A-Z0-9]{4}$/', $normalizedCn)) {
            $errorMessage = 'Invalid control number format. Format must match: PREFIX-YYYY-0000-XXXX (e.g. COMM-2026-0001-K9X2).';
        } else {
            try {
                $pdo = get_db();

                // 1. Query strictly the public_documents view
                $stmt = $pdo->prepare(
                    'SELECT control_number, category, document_title, public_status, created_at 
                     FROM public_documents 
                     WHERE control_number = ?'
                );
                $stmt->execute([$normalizedCn]);
                $document = $stmt->fetch();

                // 2. If found, query public milestones
                if ($document) {
                    $msStmt = $pdo->prepare(
                        'SELECT pm.public_status, pm.created_at 
                         FROM public_milestones pm
                         JOIN documents d ON pm.document_id = d.id
                         WHERE d.control_number = ?
                         ORDER BY pm.created_at ASC, pm.id ASC'
                    );
                    $msStmt->execute([$normalizedCn]);
                    $milestones = $msStmt->fetchAll();
                }
            } catch (Throwable $ex) {
                error_log('[Tracking Error] ' . $ex->getMessage());
                $errorMessage = 'A database error occurred while querying status. Please try again.';
            }
        }
    }
}

/**
 * Returns CSS class for status pill
 */
$activePage = 'track';

/**
 * Returns CSS class for status pill
 */
function get_status_class(string $status): string {
    return match (strtolower(str_replace(' ', '-', $status))) {
        'received' => 'portal-status-received',
        'under-review' => 'portal-status-under-review',
        'for-processing' => 'portal-status-for-processing',
        'approved' => 'portal-status-approved',
        'ready-for-release' => 'portal-status-ready-for-release',
        'rejected' => 'portal-status-rejected',
        default => 'portal-status-default'
    };
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
    <title>Track Document - OSGPortal Public Tracking</title>
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
                    <div class="portal-card-header">
                        <h1 class="portal-card-title">Track Document Status</h1>
                        <p class="portal-card-subtitle">
                            Enter the official Control Number issued on your submission receipt or stamped acknowledgment.
                        </p>
                    </div>

            <form action="track.php" method="GET" style="margin-bottom: 1.5rem;">
                <div class="portal-form-group">
                    <label class="portal-label" for="cn">External Control Number</label>
                    <div style="display: flex; gap: 0.5rem; flex-wrap: wrap;">
                        <input type="text" id="cn" name="cn" class="portal-input" style="flex: 1; min-width: 240px; font-family: var(--font-mono); font-size: 1.125rem; letter-spacing: 0.05em; text-transform: uppercase;"
                               placeholder="e.g. COMM-2026-0002-WJ8R"
                               value="<?= htmlspecialchars($searchedCn, ENT_QUOTES, 'UTF-8') ?>" required>
                        <button type="submit" class="portal-btn portal-btn-primary">Search Document</button>
                    </div>
                </div>
            </form>

            <?php if (!$searched): ?>
                <!-- Search Guidance Card for Empty State -->
                <div class="portal-search-guidance">
                    <h3>Where to Find Your Control Number</h3>
                    <p>
                        The External Control Number is a 16-character code (such as <code>COMM-2026-0002-WJ8R</code>) issued upon completing intake pre-registration.
                    </p>
                    <p style="margin-top: 0.5rem;">
                        You can find it at the top of your official email verification receipt or digital confirmation slip.
                    </p>
                </div>
            <?php endif; ?>

            <?php if ($errorMessage): ?>
                <div class="portal-alert portal-alert-danger" role="alert">
                    <span class="portal-alert-title">Query Notice:</span>
                    <?= htmlspecialchars($errorMessage, ENT_QUOTES, 'UTF-8') ?>
                </div>
            <?php elseif ($searched && !$document): ?>
                <div class="portal-alert portal-alert-warning" role="alert">
                    <span class="portal-alert-title">Document Not Found:</span>
                    No registered document was found with control number <strong><?= htmlspecialchars($searchedCn, ENT_QUOTES, 'UTF-8') ?></strong>. Please verify the code on your official receipt and check for typographical errors.
                </div>
            <?php elseif ($document): ?>
                <?php
                    $statusKey = strtolower(str_replace(' ', '-', (string)$document['public_status']));
                    $custodyText = match($statusKey) {
                        'received' => 'Office of the Secretary-General &bull; Records Section',
                        'under-review' => 'Records Section &bull; Under Evaluation',
                        'for-processing' => 'Assigned Action Section &bull; In Queue',
                        'approved' => 'Office of the Secretary-General &bull; Endorsement Complete',
                        'ready-for-release' => 'Office of the Secretary-General &bull; Ready for Digital Release',
                        'rejected' => 'Closed &bull; Records Section',
                        default => 'BTA OSG Records Section'
                    };
                ?>
                <!-- Document Details Card -->
                <div class="portal-receipt">
                    <div class="portal-receipt-header">
                        <div>
                            <div style="font-size: 0.8125rem; color: var(--color-text-muted); font-weight: 700; letter-spacing: 0.05em;">OFFICIAL CONTROL NUMBER</div>
                            <div class="portal-control-number-badge" style="margin: 0.25rem 0 0 0;">
                                <?= htmlspecialchars((string)$document['control_number'], ENT_QUOTES, 'UTF-8') ?>
                            </div>
                            <div class="portal-actions-bar">
                                <button type="button" class="portal-btn-utility" id="btn_copy_cn" onclick="copyControlNumber('<?= htmlspecialchars((string)$document['control_number'], ENT_QUOTES, 'UTF-8') ?>')">
                                    <span id="copy_text">Copy Control Number</span>
                                </button>
                                <button type="button" class="portal-btn-utility" onclick="window.print()">
                                    <span>Print Tracking Receipt</span>
                                </button>
                            </div>
                        </div>
                        <div>
                            <span class="portal-status-pill <?= get_status_class((string)$document['public_status']) ?>">
                                <?= htmlspecialchars((string)$document['public_status'], ENT_QUOTES, 'UTF-8') ?>
                            </span>
                        </div>
                    </div>

                    <div class="portal-kv-list">
                        <div class="portal-kv-label">Document Title:</div>
                        <div class="portal-kv-value" style="font-weight: 600; color: var(--color-text);">
                            <?= htmlspecialchars((string)$document['document_title'], ENT_QUOTES, 'UTF-8') ?>
                        </div>

                        <div class="portal-kv-label">Classification Category:</div>
                        <div class="portal-kv-value">
                            <?= htmlspecialchars((string)$document['category'], ENT_QUOTES, 'UTF-8') ?>
                        </div>

                        <div class="portal-kv-label">Registered Date:</div>
                        <div class="portal-kv-value">
                            <?= format_pst_time((string)$document['created_at']) ?>
                        </div>

                        <div class="portal-kv-label">Current Custody:</div>
                        <div class="portal-kv-value" style="color: var(--color-accent); font-weight: 600;">
                            <?= $custodyText ?>
                        </div>
                    </div>

                    <?php if ($statusKey === 'rejected'): ?>
                    <div class="portal-alert portal-alert-danger" role="alert">
                        <span class="portal-alert-title">Submission Not Approved:</span>
                        This submission was not approved for further processing. Please contact the Records Section for assistance.
                    </div>
                    <?php endif; ?>

                    <!-- Chronological Milestones History (Complete data preservation) -->
                    <div style="margin-top: 2rem;">
                        <h2 style="font-size: 1rem; color: var(--color-primary); margin-bottom: 1rem; border-bottom: 1px solid var(--color-border); padding-bottom: 0.5rem; text-transform: uppercase; letter-spacing: 0.05em; font-weight: 700;">
                            Milestone Activity Log
                        </h2>

                        <?php if (empty($milestones)): ?>
                            <p style="color: var(--color-text-muted); font-size: 0.875rem;">No additional status transitions recorded yet.</p>
                        <?php else: ?>
                            <div class="portal-timeline">
                                <?php foreach ($milestones as $ms): ?>
                                    <div class="portal-timeline-item">
                                        <div class="portal-timeline-bullet"></div>
                                        <div class="portal-timeline-time">
                                            <?= format_pst_time((string)$ms['created_at']) ?>
                                        </div>
                                        <div class="portal-timeline-content">
                                            Status updated to
                                            <span class="portal-status-pill <?= get_status_class((string)$ms['public_status']) ?>" style="font-size: 0.75rem; padding: 0.15rem 0.6rem;">
                                                <?= htmlspecialchars((string)$ms['public_status'], ENT_QUOTES, 'UTF-8') ?>
                                            </span>
                                        </div>
                                    </div>
                                <?php endforeach; ?>
                            </div>
                        <?php endif; ?>
                    </div>
                </div>
            <?php endif; ?>
        </section>
    </div>

    <aside class="portal-layout-sidebar" aria-label="Tracking Guide and Information">
        <div class="portal-sidebar-card">
            <div class="portal-sidebar-title">
                <span>Control Number Guide</span>
            </div>
            <p class="portal-sidebar-text">
                Control numbers follow the official four-part format: <code>PREFIX-YYYY-0000-XXXX</code>.
            </p>
            <ul class="portal-sidebar-list">
                <li>
                    <span class="portal-sidebar-badge">COMM</span>
                    <span>General Communications and Endorsements</span>
                </li>
                <li>
                    <span class="portal-sidebar-badge">LEG</span>
                    <span>Legislative Measures and Resolutions</span>
                </li>
                <li>
                    <span class="portal-sidebar-badge">FIN</span>
                    <span>Financial and Budgetary Documents</span>
                </li>
                <li>
                    <span class="portal-sidebar-badge">TO</span>
                    <span>Official Travel Orders and Itineraries</span>
                </li>
            </ul>
        </div>

        <div class="portal-sidebar-card">
            <div class="portal-sidebar-title">
                <span>Status Key</span>
            </div>
            <div class="portal-info-item">
                <span class="portal-info-label">Received:</span>
                Document digitally registered and verified by the Records Section.
            </div>
            <div class="portal-info-item">
                <span class="portal-info-label">Under Review:</span>
                Subject to substantive review and verification by assigned section.
            </div>
            <div class="portal-info-item">
                <span class="portal-info-label">For Processing:</span>
                Routed to the appropriate action section for processing.
            </div>
            <div class="portal-info-item">
                <span class="portal-info-label">Approved:</span>
                Executive endorsement or sign-off by the Office of the Secretary-General.
            </div>
            <div class="portal-info-item">
                <span class="portal-info-label">Ready for Release:</span>
                Action completed; approved document released digitally to requesting party.
            </div>
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
    function copyControlNumber(text) {
        if (navigator.clipboard && navigator.clipboard.writeText) {
            navigator.clipboard.writeText(text).then(function() {
                var btn = document.getElementById('copy_text');
                if (btn) {
                    var prev = btn.textContent;
                    btn.textContent = 'Copied to Clipboard!';
                    setTimeout(function() { btn.textContent = prev; }, 2500);
                }
            });
        } else {
            prompt('Copy Control Number:', text);
        }
    }

    document.addEventListener('DOMContentLoaded', function() {
        var cnInput = document.getElementById('cn');
        if (!cnInput) return;

        function formatControlNumber(val) {
            var clean = val.replace(/[^A-Za-z0-9]/g, '').toUpperCase();
            if (!clean) return '';

            var prefix = '';
            var rest = '';

            if (clean.indexOf('COMM') === 0) {
                prefix = 'COMM';
                rest = clean.slice(4);
            } else if (clean.indexOf('LEG') === 0 || clean.indexOf('FIN') === 0) {
                prefix = clean.slice(0, 3);
                rest = clean.slice(3);
            } else if (clean.indexOf('TO') === 0) {
                prefix = clean.slice(0, 2);
                rest = clean.slice(2);
            } else {
                var m = clean.match(/^([A-Z]{1,4})(.*)$/);
                if (m) {
                    prefix = m[1];
                    rest = m[2];
                } else {
                    return clean;
                }
            }

            var parts = [prefix];
            if (rest.length > 0) {
                parts.push(rest.slice(0, 4));
            }
            if (rest.length > 4) {
                parts.push(rest.slice(4, 8));
            }
            if (rest.length > 8) {
                parts.push(rest.slice(8, 12));
            }

            return parts.join('-');
        }

        cnInput.addEventListener('input', function(e) {
            var currentVal = cnInput.value;
            if (e.inputType && e.inputType.indexOf('delete') !== -1) {
                return;
            }

            var formatted = formatControlNumber(currentVal);
            if (formatted !== currentVal) {
                var selStart = cnInput.selectionStart;
                var prevLen = currentVal.length;
                cnInput.value = formatted;
                if (selStart >= prevLen) {
                    cnInput.setSelectionRange(formatted.length, formatted.length);
                }
            }
        });

        cnInput.addEventListener('blur', function() {
            if (cnInput.value) {
                cnInput.value = formatControlNumber(cnInput.value);
            }
        });

        cnInput.addEventListener('paste', function() {
            setTimeout(function() {
                cnInput.value = formatControlNumber(cnInput.value);
            }, 10);
        });
    });
    </script>
    <script src="assets/header-preview.js"></script>
</body>
</html>
