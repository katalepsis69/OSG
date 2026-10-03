<?php
// BTA OSG External Intake Portal: Bridge API Authentication Guard
// Enforces X-Bridge-Key header verification using constant-time hash_equals.

declare(strict_types=1);

require_once __DIR__ . '/db.php';

/**
 * Validates the incoming X-Bridge-Key request header against the configured
 * shared secret. Emits HTTP 401 and halts execution immediately on mismatch.
 */
function require_bridge_auth(): void {
    $headerKey = $_SERVER['HTTP_X_BRIDGE_KEY'] ?? '';
    $secret = defined('BRIDGE_SECRET_KEY') ? BRIDGE_SECRET_KEY : '';

    if (empty($secret) || str_starts_with($secret, 'REPLACE_WITH_') || empty($headerKey) || !hash_equals($secret, $headerKey)) {
        http_response_code(401);
        header('Content-Type: application/json; charset=utf-8');
        echo json_encode([
            'success' => false,
            'error' => 'Unauthorized: invalid or missing bridge authentication key.'
        ], JSON_UNESCAPED_SLASHES);
        exit;
    }
}
