<?php
// BTA OSG External Intake Portal: Session and CSRF bootstrap
// One hardened session_start for every public page: HttpOnly + SameSite cookies, a
// Secure flag whenever the request arrived over HTTPS, and a per-session CSRF token
// for the POST forms. The session id is regenerated when a pending token is first
// bound, closing the fixation window between anonymous browsing and verification.

declare(strict_types=1);

function portal_session_start(): void {
    if (session_status() === PHP_SESSION_ACTIVE) {
        return;
    }
    $isHttps = (!empty($_SERVER['HTTPS']) && $_SERVER['HTTPS'] !== 'off')
        || (($_SERVER['HTTP_X_FORWARDED_PROTO'] ?? '') === 'https');
    session_set_cookie_params([
        'lifetime' => 0,
        'path' => '/',
        'secure' => $isHttps,
        'httponly' => true,
        'samesite' => 'Lax',
    ]);
    session_start();

    // The session id is rotated once, at the moment the session first carries portal
    // state (a pending submission token), so a stolen pre-auth id cannot be replayed
    // into a verification-capable session.
    if (!empty($_SESSION['pending_token']) && empty($_SESSION['session_bound'])) {
        session_regenerate_id(true);
        $_SESSION['session_bound'] = true;
    }
}

function portal_csrf_token(): string {
    if (empty($_SESSION['csrf_token'])) {
        $_SESSION['csrf_token'] = bin2hex(random_bytes(32));
    }
    return $_SESSION['csrf_token'];
}

function portal_csrf_verify(?string $token): void {
    $expected = $_SESSION['csrf_token'] ?? '';
    if ($expected === '' || empty($token) || !hash_equals($expected, (string)$token)) {
        http_response_code(403);
        header('Content-Type: text/plain; charset=utf-8');
        exit('Request rejected: invalid or missing CSRF token. Reload the page and try again.');
    }
}
