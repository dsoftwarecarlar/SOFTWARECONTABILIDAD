<?php
declare(strict_types=1);

require __DIR__ . '/includes/app.php';

if (PHP_SAPI !== 'cli') {
    http_response_code(403);
    echo "Solo disponible por CLI." . PHP_EOL;
    exit(1);
}

function cleanup_delete_tree(string $path): void
{
    if (!is_dir($path)) {
        return;
    }

    $items = scandir($path);
    if ($items === false) {
        return;
    }

    foreach ($items as $item) {
        if ($item === '.' || $item === '..') {
            continue;
        }

        $child = app_join_path($path, $item);
        if (is_dir($child)) {
            cleanup_delete_tree($child);
            continue;
        }

        @unlink($child);
    }

    @rmdir($path);
}

function cleanup_service_prefixes(): array
{
    return [
        'servicios_changan_',
        'servicios_peug_',
        'servicios_szk_',
        'servicios_tyt_',
    ];
}

$beforeOutputs = count(glob(app_storage_path('outputs' . DIRECTORY_SEPARATOR . '*')) ?: []);
$beforeUploads = count(glob(app_storage_path('uploads' . DIRECTORY_SEPARATOR . '*')) ?: []);

app_cleanup_output_files_for_action('accion1', app_output_retention_limit());
app_cleanup_output_files_for_action('accion2', app_output_retention_limit());
app_cleanup_output_files_for_action('accion3', app_output_retention_limit());
app_cleanup_output_files_for_action('accion4', app_output_retention_limit());
app_cleanup_output_files_for_action('bundle', app_output_retention_limit());

foreach (cleanup_service_prefixes() as $prefix) {
    app_cleanup_output_files(
        app_output_retention_limit(),
        static fn(string $name): bool => str_starts_with(strtolower($name), $prefix)
    );
}

$outputDir = app_storage_path('outputs');
foreach (glob(app_join_path($outputDir, '_*')) ?: [] as $path) {
    if (is_dir($path)) {
        cleanup_delete_tree($path);
        continue;
    }

    if (is_file($path)) {
        @unlink($path);
    }
}

foreach (glob(app_join_path($outputDir, 'tmp_*')) ?: [] as $path) {
    if (is_file($path)) {
        @unlink($path);
    }
}

app_cleanup_upload_files(app_upload_retention_limit());

$afterOutputs = count(glob(app_storage_path('outputs' . DIRECTORY_SEPARATOR . '*')) ?: []);
$afterUploads = count(glob(app_storage_path('uploads' . DIRECTORY_SEPARATOR . '*')) ?: []);

$summary = [
    'output_retention' => app_output_retention_limit(),
    'upload_retention' => app_upload_retention_limit(),
    'before_outputs' => $beforeOutputs,
    'after_outputs' => $afterOutputs,
    'before_uploads' => $beforeUploads,
    'after_uploads' => $afterUploads,
    'latest_outputs' => array_map(
        static fn(array $item): string => (string)$item['name'],
        array_slice(app_output_file_entries(), 0, 12)
    ),
    'latest_uploads' => array_map(
        static fn(array $item): string => (string)$item['name'],
        array_slice(app_upload_file_entries(), 0, 12)
    ),
];

echo json_encode($summary, JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES) . PHP_EOL;
