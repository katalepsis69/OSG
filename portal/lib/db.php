<?php
// BTA OSG External Intake Portal: Database Connection Helper

// OTP and counter timestamps are stored as UTC strings and strtotime compares assume UTC regardless of the host php.ini.
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
 * Returns a configured PDO connection to the MySQL 8 portal database.
 * Sets strict error reporting, associative fetching, and UTC timezone.
 */
function get_db(): PDO {
    static $pdo = null;

    if ($pdo === null) {
        $dsn = sprintf(
            'mysql:host=%s;port=%s;dbname=%s;charset=%s',
            defined('DB_HOST') ? DB_HOST : '127.0.0.1',
            defined('DB_PORT') ? DB_PORT : '3306',
            defined('DB_NAME') ? DB_NAME : 'osg',
            defined('DB_CHARSET') ? DB_CHARSET : 'utf8mb4'
        );

        $options = [
            PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION,
            PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC,
            PDO::ATTR_EMULATE_PREPARES => false,
        ];

        $user = defined('DB_USER') ? DB_USER : 'portal';
        $pass = defined('DB_PASS') ? DB_PASS : '';

        $pdo = new PDO($dsn, $user, $pass, $options);
        $pdo->exec("SET time_zone = '+00:00'");
        ensure_portal_schema($pdo);
    }

    return $pdo;
}

function ensure_portal_schema(PDO $pdo): void {
    static $checked = false;
    if ($checked) return;
    $checked = true;

    try {
        $cols = $pdo->query("SHOW COLUMNS FROM pending_submissions LIKE 'gender'")->fetchAll();
        if (empty($cols)) {
            $pdo->exec("ALTER TABLE pending_submissions ADD COLUMN gender VARCHAR(30) NULL AFTER phone");
        }
    } catch (Throwable $e) {
        // Ignore if already migrated or table pending setup
    }

    try {
        $cols = $pdo->query("SHOW COLUMNS FROM requesters LIKE 'gender'")->fetchAll();
        if (empty($cols)) {
            $pdo->exec("ALTER TABLE requesters ADD COLUMN gender VARCHAR(30) NULL AFTER phone");
        }
    } catch (Throwable $e) {
    }
}
