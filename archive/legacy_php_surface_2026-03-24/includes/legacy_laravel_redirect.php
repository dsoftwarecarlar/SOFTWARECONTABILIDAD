<?php
declare(strict_types=1);

require_once __DIR__ . '/app.php';

function app_should_redirect_legacy_request(array $methods = ['GET', 'HEAD']): bool
{
    if (PHP_SAPI === 'cli' || PHP_SAPI === 'phpdbg') {
        return false;
    }

    $method = strtoupper((string) ($_SERVER['REQUEST_METHOD'] ?? 'GET'));
    return in_array($method, $methods, true);
}

function app_redirect_legacy_to_laravel(
    string $relative,
    array $methods = ['GET', 'HEAD'],
    bool $preserveQuery = false,
    int $status = 302
): void {
    if (!app_should_redirect_legacy_request($methods)) {
        return;
    }

    $target = app_laravel_url($relative);
    $query = trim((string) ($_SERVER['QUERY_STRING'] ?? ''));
    if ($preserveQuery && $query !== '') {
        $target .= (str_contains($target, '?') ? '&' : '?') . $query;
    }

    header('Cache-Control: no-store, private');
    header('X-Legacy-Bridge: laravel');
    header('Location: ' . $target, true, $status);
    exit;
}

function app_redirect_legacy_download_to_laravel(?string $fileName, array $methods = ['GET', 'HEAD'], int $status = 302): void
{
    $fileName = basename((string) $fileName);
    if ($fileName === '' || !app_output_has_supported_extension($fileName)) {
        return;
    }

    app_redirect_legacy_to_laravel('downloads/' . rawurlencode($fileName), $methods, false, $status);
}
