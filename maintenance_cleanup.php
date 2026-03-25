<?php
declare(strict_types=1);

const CLEANUP_OUTPUT_RETENTION_LIMIT = 3;
const CLEANUP_UPLOAD_RETENTION_LIMIT = 20;

if (PHP_SAPI !== 'cli') {
    http_response_code(403);
    echo "Solo disponible por CLI." . PHP_EOL;
    exit(1);
}

function cleanup_root(): string
{
    return __DIR__;
}

function cleanup_join_path(string ...$parts): string
{
    $clean = [];
    foreach ($parts as $index => $part) {
        $normalized = str_replace(['/', '\\'], DIRECTORY_SEPARATOR, $part);
        if ($index === 0) {
            $clean[] = rtrim($normalized, DIRECTORY_SEPARATOR);
            continue;
        }

        $clean[] = trim($normalized, DIRECTORY_SEPARATOR);
    }

    return implode(DIRECTORY_SEPARATOR, array_filter($clean, static fn(string $part): bool => $part !== ''));
}

function cleanup_storage_path(string $relative = ''): string
{
    $base = cleanup_join_path(cleanup_root(), 'storage');
    if ($relative === '') {
        return $base;
    }

    return cleanup_join_path($base, $relative);
}

function cleanup_action_export_config_path(): string
{
    return cleanup_join_path(cleanup_root(), 'config', 'cxp', 'action_exports.json');
}

function cleanup_output_extensions(): array
{
    return ['xlsx', 'xls'];
}

function cleanup_output_retention_limit(): int
{
    return CLEANUP_OUTPUT_RETENTION_LIMIT;
}

function cleanup_upload_retention_limit(): int
{
    return CLEANUP_UPLOAD_RETENTION_LIMIT;
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

        $child = cleanup_join_path($path, $item);
        if (is_dir($child)) {
            cleanup_delete_tree($child);
            continue;
        }

        @unlink($child);
    }

    @rmdir($path);
}

function cleanup_action_export_definitions(): array
{
    static $cache = null;
    if ($cache !== null) {
        return $cache;
    }

    $path = cleanup_action_export_config_path();
    if (!is_file($path)) {
        throw new RuntimeException('No existe config/cxp/action_exports.json en el proyecto.');
    }

    try {
        $decoded = json_decode((string) file_get_contents($path), true, 512, JSON_THROW_ON_ERROR);
    } catch (Throwable $exception) {
        throw new RuntimeException('No se pudo leer config/cxp/action_exports.json: ' . $exception->getMessage(), 0, $exception);
    }

    if (!is_array($decoded)) {
        throw new RuntimeException('config/cxp/action_exports.json no contiene una lista valida de acciones.');
    }

    $definitions = [];
    foreach ($decoded as $item) {
        if (!is_array($item)) {
            continue;
        }

        $key = trim((string) ($item['key'] ?? ''));
        if ($key === '') {
            continue;
        }

        $definitions[$key] = [
            'key' => $key,
            'label' => (string) ($item['label'] ?? $key),
            'sheet_name' => (string) ($item['sheet_name'] ?? strtoupper($key)),
            'module_path' => trim((string) ($item['module_path'] ?? ''), '/'),
            'file_match' => is_array($item['file_match'] ?? null) ? $item['file_match'] : [],
        ];
    }

    $cache = $definitions;
    return $cache;
}

function cleanup_output_matches_export_definition(string $fileName, array $definition): bool
{
    $rule = is_array($definition['file_match'] ?? null) ? $definition['file_match'] : [];
    $type = strtolower(trim((string) ($rule['type'] ?? '')));
    $value = (string) ($rule['value'] ?? '');

    if ($type === 'contains') {
        return $value !== '' && stripos($fileName, $value) !== false;
    }

    if ($type === 'regex') {
        if ($value === '') {
            return false;
        }

        $flags = preg_replace('/[^imsxuADSUXJu]/', '', (string) ($rule['flags'] ?? '')) ?: '';
        $pattern = '~' . str_replace('~', '\\~', $value) . '~' . $flags;
        return preg_match($pattern, $fileName) === 1;
    }

    return false;
}

function cleanup_output_has_supported_extension(string $fileName): bool
{
    $extension = strtolower(pathinfo($fileName, PATHINFO_EXTENSION));
    return in_array($extension, cleanup_output_extensions(), true);
}

function cleanup_output_matches_action(string $fileName, string $actionKey): bool
{
    $name = strtolower(trim($fileName));
    if (!cleanup_output_has_supported_extension($name)) {
        return false;
    }

    $definitions = cleanup_action_export_definitions();
    if (isset($definitions[$actionKey])) {
        return cleanup_output_matches_export_definition($name, $definitions[$actionKey]);
    }

    return match ($actionKey) {
        'bundle' => str_contains($name, 'acciones_resumen'),
        'servicios' => str_starts_with($name, 'servicios_'),
        'repuestos_tytserv' => str_starts_with($name, 'repuestos_tytserv_'),
        default => false,
    };
}

function cleanup_output_storage_files(): array
{
    $files = [];
    foreach (cleanup_output_extensions() as $extension) {
        $pattern = cleanup_storage_path('outputs' . DIRECTORY_SEPARATOR . '*.' . $extension);
        foreach (glob($pattern) ?: [] as $file) {
            if (is_file($file)) {
                $files[] = $file;
            }
        }
    }

    return array_values(array_unique($files));
}

function cleanup_output_file_timestamp(string $file): int
{
    if (!is_file($file)) {
        return 0;
    }

    $created = @filectime($file);
    if (is_int($created) && $created > 0) {
        return $created;
    }

    $modified = @filemtime($file);
    if (is_int($modified) && $modified > 0) {
        return $modified;
    }

    return time();
}

function cleanup_output_file_entries(): array
{
    $files = cleanup_output_storage_files();
    usort(
        $files,
        static function (string $a, string $b): int {
            $timeA = cleanup_output_file_timestamp($a);
            $timeB = cleanup_output_file_timestamp($b);
            if ($timeA === $timeB) {
                return strcmp(basename($b), basename($a));
            }

            return $timeB <=> $timeA;
        }
    );

    $items = [];
    foreach ($files as $file) {
        if (!is_file($file)) {
            continue;
        }

        $timestamp = cleanup_output_file_timestamp($file);
        $items[] = [
            'path' => $file,
            'name' => basename($file),
            'size' => filesize($file) ?: 0,
            'timestamp' => $timestamp,
            'time' => date('Y-m-d H:i:s', $timestamp),
        ];
    }

    return $items;
}

function cleanup_upload_file_entries(): array
{
    $entries = [];
    foreach (glob(cleanup_storage_path('uploads' . DIRECTORY_SEPARATOR . '*')) ?: [] as $file) {
        if (!is_file($file)) {
            continue;
        }

        $timestamp = filectime($file) ?: (filemtime($file) ?: time());
        $entries[] = [
            'path' => $file,
            'name' => basename($file),
            'size' => filesize($file) ?: 0,
            'timestamp' => $timestamp,
            'time' => date('Y-m-d H:i:s', $timestamp),
        ];
    }

    usort(
        $entries,
        static function (array $a, array $b): int {
            if ($a['timestamp'] === $b['timestamp']) {
                return strcmp($b['name'], $a['name']);
            }

            return $b['timestamp'] <=> $a['timestamp'];
        }
    );

    return $entries;
}

function cleanup_output_files(int $keep, ?callable $filter = null): void
{
    if ($keep < 0) {
        $keep = 0;
    }

    $entries = [];
    foreach (cleanup_output_storage_files() as $file) {
        $name = basename($file);
        if ($filter !== null && !$filter($name)) {
            continue;
        }

        $entries[] = [
            'path' => $file,
            'name' => $name,
            'time' => cleanup_output_file_timestamp($file),
        ];
    }

    usort(
        $entries,
        static function (array $a, array $b): int {
            if ($a['time'] === $b['time']) {
                return strcmp($b['name'], $a['name']);
            }

            return $b['time'] <=> $a['time'];
        }
    );

    foreach (array_slice($entries, $keep) as $entry) {
        $excelPath = (string) $entry['path'];
        if (is_file($excelPath)) {
            @unlink($excelPath);
        }

        $baseName = pathinfo((string) $entry['name'], PATHINFO_FILENAME);
        $auditPath = cleanup_join_path(cleanup_storage_path('outputs'), $baseName . '_auditoria.json');
        if (is_file($auditPath)) {
            @unlink($auditPath);
        }
    }
}

function cleanup_output_files_for_action(string $actionKey, int $keep): void
{
    cleanup_output_files(
        $keep,
        static fn(string $name): bool => cleanup_output_matches_action($name, $actionKey)
    );
}

function cleanup_upload_files(int $keep): void
{
    if ($keep < 0) {
        $keep = 0;
    }

    foreach (array_slice(cleanup_upload_file_entries(), $keep) as $entry) {
        $path = (string) $entry['path'];
        if (is_file($path)) {
            @unlink($path);
        }
    }
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

$beforeOutputs = count(glob(cleanup_storage_path('outputs' . DIRECTORY_SEPARATOR . '*')) ?: []);
$beforeUploads = count(glob(cleanup_storage_path('uploads' . DIRECTORY_SEPARATOR . '*')) ?: []);

cleanup_output_files_for_action('accion1', cleanup_output_retention_limit());
cleanup_output_files_for_action('accion2', cleanup_output_retention_limit());
cleanup_output_files_for_action('accion3', cleanup_output_retention_limit());
cleanup_output_files_for_action('accion4', cleanup_output_retention_limit());
cleanup_output_files_for_action('bundle', cleanup_output_retention_limit());

foreach (cleanup_service_prefixes() as $prefix) {
    cleanup_output_files(
        cleanup_output_retention_limit(),
        static fn(string $name): bool => str_starts_with(strtolower($name), $prefix)
    );
}

$outputDir = cleanup_storage_path('outputs');
foreach (glob(cleanup_join_path($outputDir, '_*')) ?: [] as $path) {
    if (is_dir($path)) {
        cleanup_delete_tree($path);
        continue;
    }

    if (is_file($path)) {
        @unlink($path);
    }
}

foreach (glob(cleanup_join_path($outputDir, 'tmp_*')) ?: [] as $path) {
    if (is_file($path)) {
        @unlink($path);
    }
}

cleanup_upload_files(cleanup_upload_retention_limit());

$afterOutputs = count(glob(cleanup_storage_path('outputs' . DIRECTORY_SEPARATOR . '*')) ?: []);
$afterUploads = count(glob(cleanup_storage_path('uploads' . DIRECTORY_SEPARATOR . '*')) ?: []);

$summary = [
    'output_retention' => cleanup_output_retention_limit(),
    'upload_retention' => cleanup_upload_retention_limit(),
    'before_outputs' => $beforeOutputs,
    'after_outputs' => $afterOutputs,
    'before_uploads' => $beforeUploads,
    'after_uploads' => $afterUploads,
    'latest_outputs' => array_map(
        static fn(array $item): string => (string) $item['name'],
        array_slice(cleanup_output_file_entries(), 0, 12)
    ),
    'latest_uploads' => array_map(
        static fn(array $item): string => (string) $item['name'],
        array_slice(cleanup_upload_file_entries(), 0, 12)
    ),
];

echo json_encode($summary, JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES) . PHP_EOL;
