<?php
// BTA OSG External Intake Portal: Brevo HTTP Mail Client
// Dispatches emails over HTTPS port 443 via Brevo REST API v3.
// Works reliably on cloud VPS tiers where outbound SMTP ports are blocked.
// Zero third-party dependencies (native cURL only).

// Included standalone as well: strtotime/date use UTC regardless of the host php.ini.
date_default_timezone_set('UTC');

$configPath = __DIR__ . '/../config/config.php';
if (!file_exists($configPath)) {
    error_log('config.php missing');
    if (PHP_SAPI !== 'cli') {
        http_response_code(500);
    }
    exit('Server configuration error');
}
require_once $configPath;

/**
 * Sends a transactional email using Brevo's HTTP API over port 443.
 *
 * @param string $toEmail Recipient email address
 * @param string $toName Recipient display name
 * @param string $subject Email subject line
 * @param string $htmlContent HTML formatted body
 * @return array ['success' => bool, 'message_id' => ?string, 'error' => ?string]
 */
function send_brevo_mail(string $toEmail, string $toName, string $subject, string $htmlContent): array {
    $apiKey = defined('BREVO_API_KEY') ? BREVO_API_KEY : '';
    $isPlaceholder = empty($apiKey)
        || preg_match('/^xkeysib-(REPLACE|YOUR[_\-]|CHANGEME|PASTE|INSERT|TODO|XXX)/i', $apiKey) === 1
        || strlen($apiKey) < 20;
    // Recipient addresses are PII: logs carry the masked form only.
    $maskedRecipient = preg_replace('/^(.).*(@.*)$/', '$1***$2', $toEmail);
    if ($isPlaceholder) {
        error_log('[Mail] Brevo API key is not configured. Email to ' . $maskedRecipient . ' skipped.');
        return [
            'success' => false,
            'message_id' => null,
            'error' => 'Brevo API key is not configured. Set BREVO_API_KEY in portal/config/config.php'
        ];
    }

    $senderEmail = defined('BREVO_SENDER_EMAIL') && !empty(BREVO_SENDER_EMAIL) ? BREVO_SENDER_EMAIL : '';
    if ($senderEmail === '') {
        error_log('[Mail] Brevo sender identity is not configured. Email to ' . $maskedRecipient . ' skipped.');
        return [
            'success' => false,
            'message_id' => null,
            'error' => 'Brevo sender identity is not configured. Set BREVO_SENDER_EMAIL in portal/config/config.php'
        ];
    }
    $senderName = defined('BREVO_SENDER_NAME') && !empty(BREVO_SENDER_NAME) ? BREVO_SENDER_NAME : 'BTA OSG Records Section';

    $payload = [
        'sender' => [
            'name' => $senderName,
            'email' => $senderEmail
        ],
        'to' => [
            [
                'email' => $toEmail,
                'name' => $toName
            ]
        ],
        'subject' => $subject,
        'htmlContent' => $htmlContent
    ];

    $ch = curl_init('https://api.brevo.com/v3/smtp/email');
    curl_setopt_array($ch, [
        CURLOPT_POST => true,
        CURLOPT_POSTFIELDS => json_encode($payload, JSON_UNESCAPED_SLASHES | JSON_UNESCAPED_UNICODE),
        CURLOPT_RETURNTRANSFER => true,
        CURLOPT_HTTPHEADER => [
            'api-key: ' . $apiKey,
            'Content-Type: application/json',
            'Accept: application/json'
        ],
        CURLOPT_TIMEOUT => 15,
        CURLOPT_CONNECTTIMEOUT => 10,
        CURLOPT_SSL_VERIFYPEER => true,
        CURLOPT_SSL_VERIFYHOST => 2
    ]);

    $caCandidates = [
        __DIR__ . '/../runtime/php/cacert.pem',
        __DIR__ . '/../cacert.pem',
        'C:\Program Files\Git\mingw64\etc\ssl\certs\ca-bundle.crt'
    ];
    foreach ($caCandidates as $ca) {
        if (file_exists($ca)) {
            curl_setopt($ch, CURLOPT_CAINFO, $ca);
            break;
        }
    }

    $response = curl_exec($ch);
    $httpCode = curl_getinfo($ch, CURLINFO_HTTP_CODE);
    $curlError = curl_error($ch);
    curl_close($ch);

    if ($response === false || !empty($curlError)) {
        error_log('[Mail Error] cURL failure sending to ' . $maskedRecipient . ': ' . $curlError);
        return [
            'success' => false,
            'message_id' => null,
            'error' => $curlError ?: 'Unknown cURL error'
        ];
    }

    $responseData = json_decode($response, true);
    if ($httpCode >= 200 && $httpCode < 300) {
        $messageId = $responseData['messageId'] ?? null;
        return [
            'success' => true,
            'message_id' => $messageId,
            'error' => null
        ];
    }

    $errMsg = $responseData['message'] ?? ('HTTP ' . $httpCode . ': ' . $response);
    error_log('[Mail Error] Brevo API responded with error: ' . $errMsg);
    return [
        'success' => false,
        'message_id' => null,
        'error' => $errMsg
    ];
}
