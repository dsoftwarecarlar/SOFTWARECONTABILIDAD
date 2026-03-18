<?php
declare(strict_types=1);

date_default_timezone_set('America/Guayaquil');

const APP_OUTPUT_RETENTION_LIMIT = 3;
const APP_UPLOAD_RETENTION_LIMIT = 20;
const APP_PUBLIC_HOST = '192.168.100.182';

function app_root(): string
{
    return dirname(__DIR__);
}

function app_join_path(string ...$parts): string
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

    return implode(DIRECTORY_SEPARATOR, array_filter($clean, static fn($part) => $part !== ''));
}

function app_first_existing_path(string ...$candidates): string
{
    foreach ($candidates as $candidate) {
        if ($candidate !== '' && file_exists($candidate)) {
            return $candidate;
        }
    }

    return $candidates[0] ?? '';
}

function app_storage_path(string $relative = ''): string
{
    $base = app_join_path(app_root(), 'storage');
    if ($relative === '') {
        return $base;
    }
    return app_join_path($base, $relative);
}

function app_ensure_dir(string $path): string
{
    if (!is_dir($path)) {
        mkdir($path, 0775, true);
    }

    return $path;
}

function app_base_url(): string
{
    $root = str_replace('\\', '/', realpath(app_root()) ?: app_root());
    $documentRoot = str_replace('\\', '/', realpath($_SERVER['DOCUMENT_ROOT'] ?? '') ?: ($_SERVER['DOCUMENT_ROOT'] ?? ''));
    if ($documentRoot === '' || !str_starts_with($root, $documentRoot)) {
        return '';
    }

    $suffix = trim(substr($root, strlen($documentRoot)), '/');
    return $suffix === '' ? '' : '/' . $suffix;
}

function app_request_scheme(): string
{
    $https = strtolower((string)($_SERVER['HTTPS'] ?? ''));
    if ($https === 'on' || $https === '1') {
        return 'https';
    }

    $forwardedProto = strtolower((string)($_SERVER['HTTP_X_FORWARDED_PROTO'] ?? ''));
    if ($forwardedProto === 'https') {
        return 'https';
    }

    return 'http';
}

function app_url(string $relative = ''): string
{
    $base = app_base_url();
    $relative = trim(str_replace('\\', '/', $relative), '/');
    if ($relative === '') {
        return $base === '' ? '/' : $base;
    }

    return $base === '' ? '/' . $relative : $base . '/' . $relative;
}

function app_public_root_url(): string
{
    $host = trim(APP_PUBLIC_HOST);
    if ($host === '') {
        return app_url();
    }

    return app_request_scheme() . '://' . $host . app_url();
}

function app_public_url(string $relative = ''): string
{
    $host = trim(APP_PUBLIC_HOST);
    if ($host === '') {
        return app_url($relative);
    }

    return app_request_scheme() . '://' . $host . app_url($relative);
}

function app_asset_url(string $relative = ''): string
{
    $relative = trim(str_replace('\\', '/', $relative), '/');
    if ($relative === '') {
        return app_url('assets');
    }

    return app_url('assets/' . $relative);
}

function app_output_download_url(string $fileName): string
{
    return app_url('download.php?file=' . rawurlencode($fileName));
}

function app_output_extensions(): array
{
    return ['xlsx', 'xls'];
}

function app_output_has_supported_extension(string $fileName): bool
{
    $extension = strtolower(pathinfo($fileName, PATHINFO_EXTENSION));
    return in_array($extension, app_output_extensions(), true);
}

function app_output_retention_limit(): int
{
    return APP_OUTPUT_RETENTION_LIMIT;
}

function app_upload_retention_limit(): int
{
    return APP_UPLOAD_RETENTION_LIMIT;
}

function app_logo_path(): string
{
    return app_join_path(app_root(), 'images', 'logo2.jpg');
}

function app_logo_url(): string
{
    return app_url('images/logo2.jpg');
}

function app_brand(): array
{
    return [
        'company' => 'Departamento Contable',
        'platform' => 'Portal Operativo',
        'tagline' => 'Sistema interno en produccion para procesos contables',
        'logo_url' => app_logo_url(),
        'logo_exists' => is_file(app_logo_path()),
    ];
}

function app_workspaces(): array
{
    $cxpModules = [
        [
            'slug' => 'cxp_pdf',
            'title' => 'Libro Compras Proveedores',
            'status' => 'Operativo',
            'description' => 'Lee el PDF de proveedores, clasifica los datos y genera el Excel final del libro de compras.',
            'url' => app_public_url('modules/cxp_pdf/index.php'),
            'state' => 'active',
        ],
        [
            'slug' => 'cxp_txt',
            'title' => 'Retenciones Proveedores',
            'status' => 'Operativo',
            'description' => 'Ordena el TXT de retenciones y lo lleva al Excel operativo con el formato del area.',
            'url' => app_public_url('modules/cxp_txt/index.php'),
            'state' => 'active',
        ],
        [
            'slug' => 'cxp_accion3',
            'title' => 'Mayor Retenciones',
            'status' => 'Operativo',
            'description' => 'Transforma el TXT del mayor de retenciones al Excel de control del proceso.',
            'url' => app_public_url('modules/cxp_accion3/index.php'),
            'state' => 'active',
        ],
        [
            'slug' => 'cxp_accion4',
            'title' => 'Mayor IVA',
            'status' => 'Operativo',
            'description' => 'Transforma el TXT del mayor IVA al Excel de control del proceso.',
            'url' => app_public_url('modules/cxp_accion4/index.php'),
            'state' => 'active',
        ],
        [
            'slug' => 'cxp_servicios_marcas',
            'title' => 'Conciliacion Servicios por Marca',
            'status' => 'Operativo',
            'description' => 'Separa el reporte mensual de servicios por marca y genera la plantilla operativa de cada empresa.',
            'url' => app_public_url('modules/cxp_servicios_marcas/index.php'),
            'state' => 'active',
        ],
        [
            'slug' => 'cxp_repuestos_tytserv',
            'title' => 'Facturacion Repuestos TYTSERV',
            'status' => 'Operativo',
            'description' => 'Carga 4 reportes de ventas por marca y genera la plantilla mensual consolidada de facturacion de repuestos.',
            'url' => app_public_url('modules/cxp_repuestos_tytserv/index.php'),
            'state' => 'active',
        ],
    ];

    $cxpModulesBySlug = [];
    foreach ($cxpModules as $module) {
        $cxpModulesBySlug[$module['slug']] = $module;
    }

    return [
        [
            'slug' => 'cxp',
            'title' => 'Contabilidad Talleres',
            'person' => 'Responsable asignado por contabilidad',
            'summary' => 'Proceso en produccion para leer archivos PDF, TXT y Excel, clasificar datos y generar archivos finales para carga contable.',
            'future_note' => 'La expansion del area queda organizada por ventanas operativas, sin mezclar funciones entre responsables.',
            'home_url' => app_public_url('areas/cxp/index.php'),
            'modules' => $cxpModules,
            'windows' => [
                [
                    'slug' => 'libro_compras_aclt',
                    'title' => 'Libro de Compras ACLT',
                    'summary' => 'Ventana operativa del responsable de Cuentas por Pagar para ejecutar las 4 acciones del proceso y el consolidado general.',
                    'route_note' => 'Ruta de salida: storage/outputs. Descarga publica: solo archivo .xlsx. Auditoria: archivo interno para control del area.',
                    'url' => app_public_url('areas/cxp/libro-compras-aclt.php'),
                    'modules' => [
                        $cxpModulesBySlug['cxp_pdf'],
                        $cxpModulesBySlug['cxp_txt'],
                        $cxpModulesBySlug['cxp_accion3'],
                        $cxpModulesBySlug['cxp_accion4'],
                    ],
                    'bundle' => [
                        'slug' => 'cxp_bundle',
                        'title' => 'Consolidado General',
                        'status' => 'Operativo',
                        'description' => 'Consolida las ultimas salidas validadas de las 4 acciones en un solo Excel de control.',
                        'url' => app_public_url('export_all_actions.php'),
                        'state' => 'active',
                    ],
                ],
                [
                    'slug' => 'conciliacion_servicios_marcas',
                    'title' => 'Conciliacion Servicios por Marca',
                    'summary' => 'Ventana operativa para separar el reporte mensual de servicios por marca y descargar cada plantilla final por empresa.',
                    'route_note' => 'Ruta de salida: storage/outputs. Descarga publica: archivos .xls y .xlsx. Base operativa: plantillas mensuales del area.',
                    'url' => app_public_url('areas/cxp/conciliacion-servicios-marcas.php'),
                    'modules' => [
                        $cxpModulesBySlug['cxp_servicios_marcas'],
                    ],
                ],
                [
                    'slug' => 'facturacion_repuestos_tytserv',
                    'title' => 'Facturacion Repuestos TYTSERV',
                    'summary' => 'Ventana operativa para cargar las 4 bases RepLibroVentasGeneral del mes y obtener el reporte final segun la plantilla manual.',
                    'route_note' => 'Ruta de salida: storage/outputs. Descarga publica: archivo .xlsx. Base operativa: resources/cxp/repuestos_tytserv/templates.',
                    'url' => app_public_url('areas/cxp/facturacion-repuestos-tytserv.php'),
                    'modules' => [
                        $cxpModulesBySlug['cxp_repuestos_tytserv'],
                    ],
                ],
            ],
        ],
    ];
}

function app_workspace_by_slug(string $slug): ?array
{
    foreach (app_workspaces() as $workspace) {
        if ($workspace['slug'] === $slug) {
            return $workspace;
        }
    }

    return null;
}

function app_workspace_module(string $workspaceSlug, string $moduleSlug): ?array
{
    $workspace = app_workspace_by_slug($workspaceSlug);
    if ($workspace === null) {
        return null;
    }

    foreach ($workspace['modules'] as $module) {
        if ($module['slug'] === $moduleSlug) {
            return [
                'workspace' => $workspace,
                'module' => $module,
            ];
        }
    }

    return null;
}

function app_workspace_window(string $workspaceSlug, string $windowSlug): ?array
{
    $workspace = app_workspace_by_slug($workspaceSlug);
    if ($workspace === null) {
        return null;
    }

    foreach (($workspace['windows'] ?? []) as $window) {
        if (($window['slug'] ?? '') === $windowSlug) {
            return [
                'workspace' => $workspace,
                'window' => $window,
            ];
        }
    }

    return null;
}

function app_action_export_config_path(): string
{
    return app_join_path(app_root(), 'config', 'cxp', 'action_exports.json');
}

function app_action_export_definitions(): array
{
    static $cache = null;
    if ($cache !== null) {
        return $cache;
    }

    $path = app_action_export_config_path();
    if (!is_file($path)) {
        throw new RuntimeException('No existe config/cxp/action_exports.json en el proyecto.');
    }

    try {
        $decoded = json_decode((string)file_get_contents($path), true, 512, JSON_THROW_ON_ERROR);
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

        $key = trim((string)($item['key'] ?? ''));
        if ($key === '') {
            continue;
        }

        $bundleExtensions = array_values(array_filter(
            array_map(
                static fn($extension): string => strtolower(ltrim(trim((string)$extension), '.')),
                is_array($item['bundle_extensions'] ?? null) ? $item['bundle_extensions'] : []
            ),
            static fn(string $extension): bool => $extension !== ''
        ));

        $definitions[$key] = [
            'key' => $key,
            'label' => (string)($item['label'] ?? $key),
            'sheet_name' => (string)($item['sheet_name'] ?? strtoupper($key)),
            'module_path' => trim((string)($item['module_path'] ?? ''), '/'),
            'bundle_extensions' => $bundleExtensions,
            'file_match' => is_array($item['file_match'] ?? null) ? $item['file_match'] : [],
        ];
    }

    $cache = $definitions;
    return $cache;
}

function app_output_matches_export_definition(string $fileName, array $definition): bool
{
    $rule = is_array($definition['file_match'] ?? null) ? $definition['file_match'] : [];
    $type = strtolower(trim((string)($rule['type'] ?? '')));
    $value = (string)($rule['value'] ?? '');

    if ($type === 'contains') {
        return $value !== '' && stripos($fileName, $value) !== false;
    }

    if ($type === 'regex') {
        if ($value === '') {
            return false;
        }
        $flags = preg_replace('/[^imsxuADSUXJu]/', '', (string)($rule['flags'] ?? '')) ?: '';
        $pattern = '~' . str_replace('~', '\\~', $value) . '~' . $flags;
        return preg_match($pattern, $fileName) === 1;
    }

    return false;
}

function app_output_matches_action(string $fileName, string $actionKey): bool
{
    $name = strtolower(trim($fileName));
    if (!app_output_has_supported_extension($name)) {
        return false;
    }

    $actionDefinitions = app_action_export_definitions();
    if (isset($actionDefinitions[$actionKey])) {
        return app_output_matches_export_definition($name, $actionDefinitions[$actionKey]);
    }

    return match ($actionKey) {
        'bundle' => str_contains($name, 'acciones_resumen'),
        'servicios' => str_starts_with($name, 'servicios_'),
        'repuestos_tytserv' => str_starts_with($name, 'repuestos_tytserv_'),
        default => false,
    };
}

function app_output_storage_files(): array
{
    $files = [];
    foreach (app_output_extensions() as $extension) {
        $pattern = app_storage_path('outputs' . DIRECTORY_SEPARATOR . '*.' . $extension);
        foreach (glob($pattern) ?: [] as $file) {
            $files[] = $file;
        }
    }

    return array_values(array_unique($files));
}

function app_output_file_entries(): array
{
    $files = app_output_storage_files();
    usort(
        $files,
        static function (string $a, string $b): int {
            $timeA = filectime($a) ?: (filemtime($a) ?: 0);
            $timeB = filectime($b) ?: (filemtime($b) ?: 0);
            if ($timeA === $timeB) {
                return strcmp(basename($b), basename($a));
            }
            return $timeB <=> $timeA;
        }
    );

    $items = [];
    foreach ($files as $file) {
        $timestamp = filectime($file) ?: (filemtime($file) ?: time());
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

function app_list_output_files(int $limit = 20): array
{
    return array_slice(app_output_file_entries(), 0, $limit);
}

function app_list_output_files_for_action(string $actionKey, int $limit = 20): array
{
    $items = array_values(array_filter(
        app_output_file_entries(),
        static fn(array $item): bool => app_output_matches_action((string)$item['name'], $actionKey)
    ));

    return array_slice($items, 0, $limit);
}

function app_find_latest_output_file(string $actionKey): ?array
{
    $items = app_list_output_files_for_action($actionKey, 1);
    return $items[0] ?? null;
}

function app_cleanup_output_files_for_action(string $actionKey, int $keep = APP_OUTPUT_RETENTION_LIMIT): void
{
    app_cleanup_output_files(
        $keep,
        static fn(string $name): bool => app_output_matches_action($name, $actionKey)
    );
}

function app_latest_action_exports(): array
{
    $actions = [];
    foreach (app_action_export_definitions() as $key => $meta) {
        $actions[$key] = [
            'key' => $key,
            'label' => (string)$meta['label'],
            'sheet_name' => (string)$meta['sheet_name'],
            'module_url' => app_url((string)$meta['module_path']),
            'latest' => app_find_latest_output_file($key),
        ];
    }

    return $actions;
}

/**
 * Elimina archivos Excel antiguos de storage/outputs y su auditoria JSON asociada,
 * conservando solo los mas recientes segun el filtro recibido.
 *
 * @param int $keep Cantidad de archivos a conservar.
 * @param null|callable(string):bool $filter Recibe el nombre del archivo Excel.
 */
function app_cleanup_output_files(int $keep = APP_OUTPUT_RETENTION_LIMIT, ?callable $filter = null): void
{
    if ($keep < 0) {
        $keep = 0;
    }

    $files = app_output_storage_files();
    $entries = [];

    foreach ($files as $file) {
        $name = basename($file);
        if ($filter !== null && !$filter($name)) {
            continue;
        }

        $entries[] = [
            'path' => $file,
            'name' => $name,
            'time' => filectime($file) ?: (filemtime($file) ?: 0),
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
        $excelPath = (string)$entry['path'];
        if (is_file($excelPath)) {
            @unlink($excelPath);
        }

        $baseName = pathinfo((string)$entry['name'], PATHINFO_FILENAME);
        $auditPath = app_join_path(app_storage_path('outputs'), $baseName . '_auditoria.json');
        if (is_file($auditPath)) {
            @unlink($auditPath);
        }
    }
}

function app_upload_file_entries(): array
{
    $pattern = app_storage_path('uploads' . DIRECTORY_SEPARATOR . '*');
    $entries = [];

    foreach (glob($pattern) ?: [] as $file) {
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

function app_cleanup_upload_files(int $keep = APP_UPLOAD_RETENTION_LIMIT): void
{
    if ($keep < 0) {
        $keep = 0;
    }

    foreach (array_slice(app_upload_file_entries(), $keep) as $entry) {
        $path = (string)$entry['path'];
        if (is_file($path)) {
            @unlink($path);
        }
    }
}
