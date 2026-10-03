<?php
// BTA OSG External Intake Portal: Bridge API - Import Acknowledgment
// Endpoint: POST /api/import_ack.php
// Supports single control_number or batch control_numbers array.

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

$controlNumbers = [];
if (is_array($data)) {
    if (!empty($data['control_numbers']) && is_array($data['control_numbers'])) {
        $controlNumbers = array_values(array_filter(array_map('trim', $data['control_numbers'])));
    } elseif (!empty($data['control_number'])) {
        $controlNumbers = [trim((string)$data['control_number'])];
    }
}

if (empty($controlNumbers)) {
    http_response_code(400);
    echo json_encode(['success' => false, 'error' => 'Missing control_number or control_numbers in request body.'], JSON_UNESCAPED_SLASHES);
    exit;
}

try {
    $pdo = get_db();
    $stmt = $pdo->prepare('UPDATE documents SET imported_at = NOW() WHERE control_number = ? AND imported_at IS NULL');
    $pdo->beginTransaction();
    $totalAcknowledged = 0;
    foreach ($controlNumbers as $cn) {
        $stmt->execute([$cn]);
        $totalAcknowledged += $stmt->rowCount();
    }
    $pdo->commit();

    echo json_encode([
        'success' => true,
        'count' => count($controlNumbers),
        'acknowledged' => $totalAcknowledged,
        'message' => "Successfully processed import acknowledgment for $totalAcknowledged documents."
    ], JSON_UNESCAPED_SLASHES);

} catch (Throwable $ex) {
    if (isset($pdo) && $pdo->inTransaction()) {
        $pdo->rollBack();
    }
    http_response_code(500);
    error_log('[Bridge Error] import_ack failed: ' . $ex->getMessage());
    echo json_encode([
        'success' => false,
        'error' => 'Internal database error updating import status.'
    ], JSON_UNESCAPED_SLASHES);
}
