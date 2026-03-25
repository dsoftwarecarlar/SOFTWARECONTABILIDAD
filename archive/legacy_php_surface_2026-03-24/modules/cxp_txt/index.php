<?php
declare(strict_types=1);

require dirname(__DIR__, 2) . '/includes/legacy_laravel_redirect.php';
app_redirect_legacy_to_laravel('cxp/modules/accion2');

require dirname(__DIR__, 2) . '/includes/bootstrap.php';

use App\Shared\Application\ExternalScriptModuleController;
use App\Shared\Infrastructure\ExternalCommandRunner;
use App\Shared\Support\PhpView;

$controller = new ExternalScriptModuleController(
    new ExternalCommandRunner(),
    [
        'workspace_slug' => 'cxp',
        'module_slug' => 'cxp_txt',
        'action_key' => 'accion2',
        'upload_field' => 'txt_file',
        'accepted_extensions' => ['txt'],
        'default_upload_name' => 'archivo.txt',
        'fallback_base_name' => 'archivo',
        'script_relative_path' => 'run_bot_accion2.js',
        'output_suffix' => '_accion2.xlsx',
        'ignore_console_prefixes' => [
            'Auditoria JSON:',
            'TXT leido:',
            'Excel generado (una sola hoja):',
        ],
        'generated_path_prefix' => 'Excel generado (una sola hoja):',
        'messages' => [
            'missing_file' => 'No se recibio el archivo TXT.',
            'upload_error' => 'Error al subir el archivo TXT.',
            'invalid_extension' => 'Solo se permiten archivos TXT.',
            'save_failure' => 'No se pudo guardar el TXT subido.',
            'script_missing' => 'No existe run_bot_accion2.js en el proyecto.',
            'generated_missing' => 'No se encontro el Excel generado.',
            'process_failed' => 'El proceso fallo.',
        ],
    ]
);
$viewData = $controller->handle($_SERVER, $_FILES);

$otherModules = array_values(array_filter(
    $viewData['workspace']['modules'],
    static fn(array $module): bool => $module['slug'] !== 'cxp_txt'
));

$successMessage = '';
if ($viewData['result'] !== null) {
    $successMessage = 'Archivo generado correctamente: ' . (string)$viewData['result']['excel_name'];
    if (!empty($viewData['result']['generated_at'])) {
        $successMessage .= ' | Generado: ' . (string)$viewData['result']['generated_at'];
    }
    if (!empty($viewData['result']['output_origin_note'])) {
        $successMessage .= ' | ' . (string)$viewData['result']['output_origin_note'];
    }
}

$pageConfig = [
    'top_title' => (string)$viewData['currentModule']['title'],
    'top_note' => 'Accion 2 del flujo contable mensual',
    'back_url' => app_url('modules/cxp_pdf/index.php'),
    'back_label' => 'Volver a Accion 1',
    'hero_chip' => 'Flujo TXT en produccion',
    'hero_title' => 'TXT desordenado a Excel RET PROV con formato exacto del ejemplo manual.',
    'hero_lead' => 'Este modulo toma un TXT de retenciones, normaliza filas y columnas, y genera un Excel final en hoja RET PROV siguiendo el formato de la plantilla manual ACCION2.xlsx.',
    'hero_note' => 'Salida de usuario: descarga de archivo .xlsx.',
    'hero_card_label' => 'Plantilla base',
    'hero_card_value' => 'resources/cxp/acciones/templates/ACCION2.xlsx',
    'show_logo' => true,
    'sidebar_panels' => array_values(array_filter([
        [
            'title' => 'Mapa del proceso',
            'items' => [
                ['tag' => 'Entrada', 'title' => 'Archivo .txt', 'description' => 'Cargado por el usuario del area para la accion mensual.'],
                ['tag' => 'Transformacion', 'title' => 'Normalizacion y clasificacion', 'description' => 'Se validan columnas clave: NUM RT, FECHA, TIPO, BASE, RETENCION.'],
                ['tag' => 'Salida', 'title' => 'Excel RET PROV', 'description' => 'Archivo final con formato de ejemplo manual y resumen lateral.'],
            ],
        ],
        $otherModules !== [] ? [
            'title' => 'Otras acciones del workspace',
            'type' => 'links',
            'items' => array_map(
                static fn(array $module): array => [
                    'title' => (string)$module['title'],
                    'description' => (string)$module['description'],
                    'url' => (string)$module['url'],
                ],
                $otherModules
            ),
        ] : null,
    ])),
    'upload_title' => 'Procesar archivo TXT',
    'upload_description' => 'Sube el TXT y el sistema devolvera el Excel final para esta segunda accion.',
    'form_label' => 'Archivo TXT de retenciones',
    'form_field_name' => 'txt_file',
    'form_accept' => '.txt,text/plain',
    'button_label' => 'Generar Excel',
    'download_label' => 'Descargar Excel',
    'success_message' => $successMessage,
    'stats_items' => [
        ['tag' => 'Salida publica', 'title' => 'Solo .xlsx', 'description' => 'No se expone el JSON interno en la interfaz.'],
        ['tag' => 'Historial', 'title' => 'Ordenado por creacion', 'description' => 'Ultimos archivos arriba, mas antiguos abajo.'],
        ['tag' => 'Destino de archivos', 'title' => 'storage/outputs', 'description' => 'Ubicacion central para control del equipo.'],
    ],
    'history_title' => 'Historial reciente (Accion 2)',
    'history_description' => 'Archivos generados por este modulo TXT a Excel. Se conservan solo los ultimos ' . app_output_retention_limit() . ' informes.',
    'history_empty_title' => 'Sin archivos de Accion 2 todavia',
    'history_empty_description' => 'Se mostraran aqui cuando ejecutes el primer TXT.',
];

(new PhpView())->render(
    app_join_path(app_root(), 'templates', 'cxp', 'script_module_page.php'),
    array_merge($viewData, ['pageConfig' => $pageConfig])
);
