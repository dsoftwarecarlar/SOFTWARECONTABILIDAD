<?php

declare(strict_types=1);

$requestUri = urldecode(parse_url($_SERVER['REQUEST_URI'] ?? '/', PHP_URL_PATH) ?: '/');
$publicPath = __DIR__ . '/public';
$requestedFile = $publicPath . $requestUri;

if ($requestUri !== '/' && is_file($requestedFile)) {
    $extension = strtolower(pathinfo($requestedFile, PATHINFO_EXTENSION));
    $contentTypes = [
        'css' => 'text/css; charset=utf-8',
        'js' => 'application/javascript; charset=utf-8',
        'json' => 'application/json; charset=utf-8',
        'jpg' => 'image/jpeg',
        'jpeg' => 'image/jpeg',
        'png' => 'image/png',
        'webp' => 'image/webp',
        'svg' => 'image/svg+xml',
        'ico' => 'image/x-icon',
        'woff' => 'font/woff',
        'woff2' => 'font/woff2',
    ];

    $contentType = $contentTypes[$extension] ?? (function_exists('mime_content_type') ? mime_content_type($requestedFile) : null);
    if (is_string($contentType) && $contentType !== '') {
        header('Content-Type: ' . $contentType);
    }

    header('Content-Length: ' . (string) filesize($requestedFile));
    readfile($requestedFile);
    exit;
}

require_once $publicPath . '/index.php';
