<?php
// BTA OSG External Intake Portal: Bridge API - Public Status Update
// Endpoint: POST /api/status.php
// Idempotently updates the document public status and appends a milestone.
// Secured by: X-Bridge-Key header

declare(strict_types=1);

require_once __DIR__ . '/../../lib/bridge_auth.php';

header('Content-Type: application/json; charset=utf-8');

if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
    http_response_code(405);
    echo json_encode(['success' => false, 'error' => 'Method not allowed. Use POST.'], JSON_UNESCAPED_SLASHES);
    exit;
}

require_bridge_auth();

$rawBody = file_get_contents('php://input');
$data = json_decode($rawBody, true);

if (!is_array($data) || empty($data['control_number']) || empty($data['public_status'])) {
    http_response_code(400);
    echo json_encode([
        'success' => false,
        'error' => 'Missing control_number or public_status in request body.'
    ], JSON_UNESCAPED_SLASHES);
    exit;
}

$controlNumber = trim((string)$data['control_number']);
$publicStatus = trim((string)$data['public_status']);

$validStatuses = [
    'Received',
    'Under Review',
    'For Processing',
    'Approved',
    'Ready for Release',
    'Rejected'
];

if (!in_array($publicStatus, $validStatuses, true)) {
    http_response_code(400);
    echo json_encode([
        'success' => false,
        'error' => 'Invalid public_status. Allowed values: ' . implode(', ', $validStatuses)
    ], JSON_UNESCAPED_SLASHES);
    exit;
}

try {
    $pdo = get_db();

    // 1. Locate the document
    $docStmt = $pdo->prepare('SELECT id, public_status FROM documents WHERE control_number = ?');
    $docStmt->execute([$controlNumber]);
    $doc = $docStmt->fetch();

    if (!$doc) {
        http_response_code(404);
        echo json_encode([
            'success' => false,
            'error' => 'Document not found with control number ' . $controlNumber
        ], JSON_UNESCAPED_SLASHES);
        exit;
    }

    $documentId = (int)$doc['id'];
    $currentStatus = (string)$doc['public_status'];

    // 2. Idempotent check: if status is identical, no update needed
    if ($currentStatus === $publicStatus) {
        echo json_encode([
            'success' => true,
            'control_number' => $controlNumber,
            'status' => $publicStatus,
            'message' => 'Status already matches; no milestone appended.'
        ], JSON_UNESCAPED_SLASHES);
        exit;
    }

    // 3. Status has changed: transactionally update status and append unnotified milestone
    $pdo->beginTransaction();

    $updDoc = $pdo->prepare('UPDATE documents SET public_status = ? WHERE id = ?');
    $updDoc->execute([$publicStatus, $documentId]);

    $insMs = $pdo->prepare(
        'INSERT INTO public_milestones (document_id, public_status, created_at, notified_at)
         VALUES (?, ?, NOW(), NULL)'
    );
    $insMs->execute([$documentId, $publicStatus]);

    $pdo->commit();

    echo json_encode([
        'success' => true,
        'control_number' => $controlNumber,
        'previous_status' => $currentStatus,
        'current_status' => $publicStatus,
        'message' => 'Public status updated and milestone appended.'
    ], JSON_UNESCAPED_SLASHES);

} catch (Throwable $ex) {
    if (isset($pdo) && $pdo->inTransaction()) {
        $pdo->rollBack();
    }
    http_response_code(500);
    error_log('[Bridge Error] status.php update failed: ' . $ex->getMessage());
    echo json_encode([
        'success' => false,
        'error' => 'Internal database error updating public status.'
    ], JSON_UNESCAPED_SLASHES);
}
