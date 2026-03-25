<?php
declare(strict_types=1);

require __DIR__ . '/includes/legacy_laravel_redirect.php';
app_redirect_legacy_download_to_laravel((string) ($_GET['file'] ?? ''));

require __DIR__ . '/includes/app.php';

$file = basename((string)($_GET['file'] ?? ''));
if ($file === '' || !preg_match('/^[A-Za-z0-9._-]+\.(xlsx|xls)$/i', $file)) {
    http_response_code(404);
    exit('Archivo no disponible.');
}

$path = app_storage_path(app_join_path('outputs', $file));
if (!is_file($path)) {
    http_response_code(404);
    exit('Archivo no disponible.');
}

$extension = strtolower(pathinfo($file, PATHINFO_EXTENSION));
$contentType = match ($extension) {
    'xls' => 'application/vnd.ms-excel',
    default => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
};

header('Content-Type: ' . $contentType);
header('Content-Disposition: attachment; filename="' . $file . '"');
header('Content-Length: ' . (string)(filesize($path) ?: 0));
header('Cache-Control: private, max-age=0, must-revalidate');

readfile($path);
