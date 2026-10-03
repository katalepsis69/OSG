<?php
// BTA OSG External Intake Portal: Bridge API - External Intake Queue
// Endpoint: GET /api/external_queue.php
// Returns unimported public documents pending ingestion into the internal system.
// Secured by: X-Bridge-Key header

declare(strict_types=1);

require_once __DIR__ . '/../../lib/bridge_auth.php';

header('Content-Type: application/json; charset=utf-8');

if ($_SERVER['REQUEST_METHOD'] !== 'GET') {
    http_response_code(405);
    echo json_encode(['success' => false, 'error' => 'Method not allowed. Use GET.'], JSON_UNESCAPED_SLASHES);
    exit;
}

require_bridge_auth();

try {
    $pdo = get_db();

    $stmt = $pdo->prepare(
        'SELECT 
            d.control_number,
            d.category,
            d.document_title,
            r.full_name AS requester_name,
            r.email AS requester_email,
            r.phone AS requester_phone,
            r.gender AS requester_gender,
            d.created_at
         FROM documents d
         JOIN requesters r ON d.requester_id = r.id
         WHERE d.imported_at IS NULL
         ORDER BY d.id ASC'
    );
    $stmt->execute();
    $items = $stmt->fetchAll();

    echo json_encode([
        'success' => true,
        'count' => count($items),
        'items' => $items
    ], JSON_UNESCAPED_SLASHES | JSON_UNESCAPED_UNICODE);

} catch (Throwable $ex) {
    http_response_code(500);
    error_log('[Bridge Error] external_queue query failed: ' . $ex->getMessage());
    echo json_encode([
        'success' => false,
        'error' => 'Internal database error reading external queue.'
    ], JSON_UNESCAPED_SLASHES);
}
