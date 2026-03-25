<?php
declare(strict_types=1);

require __DIR__ . '/includes/legacy_laravel_redirect.php';
app_redirect_legacy_to_laravel('cxp/modules/consolidado-acciones');

require __DIR__ . '/includes/app.php';

$latestActions = app_latest_action_exports();
$missing = array_values(array_filter(
    $latestActions,
    static fn(array $item): bool => $item['latest'] === null
));

if (count($missing) > 0) {
    http_response_code(409);
    header('Content-Type: text/plain; charset=utf-8');
    echo 'Faltan archivos generados para: ' .
        implode(', ', array_map(static fn(array $item): string => (string)$item['label'], $missing)) . '.';
    exit;
}

$outputsDir = app_ensure_dir(app_storage_path('outputs'));
$timestamp = date('Ymd_His');
$outputFileName = 'acciones_resumen_' . $timestamp . '.xlsx';
$outputPath = app_join_path($outputsDir, $outputFileName);
$scriptPath = app_join_path(app_root(), 'run_export_all_actions.js');

if (!is_file($scriptPath)) {
    http_response_code(500);
    header('Content-Type: text/plain; charset=utf-8');
    exit('No existe run_export_all_actions.js en el proyecto.');
}

$command = escapeshellarg('node') . ' ' .
    escapeshellarg($scriptPath) . ' ' .
    escapeshellarg('--output') . ' ' .
    escapeshellarg($outputPath) . ' 2>&1';

$lines = [];
$exitCode = 0;
exec($command, $lines, $exitCode);

$console = trim(implode(PHP_EOL, array_filter(array_map('trim', $lines), static fn(string $line): bool => $line !== '')));
if ($exitCode !== 0) {
    http_response_code(500);
    header('Content-Type: text/plain; charset=utf-8');
    exit($console === '' ? 'No se pudo generar el Excel consolidado.' : $console);
}

$generatedPath = $outputPath;
foreach ($lines as $line) {
    $trimmed = trim((string)$line);
    if (stripos($trimmed, 'Excel consolidado generado:') === 0) {
        $generatedPath = trim(substr($trimmed, strlen('Excel consolidado generado:')));
        break;
    }
}

$generatedReal = realpath($generatedPath) ?: $generatedPath;
if (!is_file($generatedReal)) {
    http_response_code(500);
    header('Content-Type: text/plain; charset=utf-8');
    exit('No se encontro el Excel consolidado generado.');
}

app_cleanup_output_files_for_action('bundle', app_output_retention_limit());

$downloadName = basename($generatedReal);
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment; filename="' . $downloadName . '"');
header('Content-Length: ' . (string)(filesize($generatedReal) ?: 0));
header('Cache-Control: private, max-age=0, must-revalidate');

readfile($generatedReal);
