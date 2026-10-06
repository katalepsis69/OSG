<?php
// BTA OSG External Intake Portal: shared file-backed per-IP rate limiter.
// Used by the public endpoints (track, submit, resend) so each one gets a
// per-client bucket instead of relying on the Nginx zone that covers only
// /track.php.

declare(strict_types=1);

function osg_rate_limit_dir(): string {
    $dir = sys_get_temp_dir() . DIRECTORY_SEPARATOR . 'bta_rates';
    if (!is_dir($dir)) {
        @mkdir($dir, 0755, true);
    }
    return $dir;
}

/**
 * The bucketing client identity. CF-Connecting-IP is honoured only when the
 * request arrives over loopback: the cloudflared tunnel connects from
 * 127.0.0.1 and Cloudflare overwrites the header, so behind the tunnel it is
 * genuine. In the cloud posture Nginx delivers a real client REMOTE_ADDR, and
 * a rotated forged header buys an attacker nothing there.
 */
function osg_rate_client_ip(): string {
    $remote = $_SERVER['REMOTE_ADDR'] ?? '127.0.0.1';
    if (in_array($remote, ['127.0.0.1', '::1'], true) && !empty($_SERVER['HTTP_CF_CONNECTING_IP'])) {
        return trim((string)$_SERVER['HTTP_CF_CONNECTING_IP']);
    }
    return (string)$remote;
}

/**
 * Records one hit and reports whether the caller is still under the cap.
 */
function osg_rate_limit(string $bucket, int $max, int $windowSeconds): bool {
    $file = osg_rate_limit_dir() . DIRECTORY_SEPARATOR . $bucket . '_' . md5(osg_rate_client_ip()) . '.json';
    $now = time();
    $requests = [];

    $fp = fopen($file, 'c+');
    if (!$fp) {
        return true;
    }

    // One exclusive lock covers read, prune, and append: a shared-lock read let
    // concurrent requests all observe the pre-increment count and slip past the cap.
    $allowed = true;
    if (flock($fp, LOCK_EX)) {
        $decoded = json_decode((string)stream_get_contents($fp), true);
        if (is_array($decoded)) {
            $requests = array_values(array_filter($decoded, fn($ts) => ($now - (int)$ts) < $windowSeconds));
        }
        if (count($requests) < $max) {
            $requests[] = $now;
            ftruncate($fp, 0);
            rewind($fp);
            fwrite($fp, json_encode($requests));
        } else {
            $allowed = false;
        }
        flock($fp, LOCK_UN);
    }
    fclose($fp);
    return $allowed;
}
