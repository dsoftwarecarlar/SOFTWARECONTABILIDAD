<?php
declare(strict_types=1);

require dirname(__DIR__, 2) . '/includes/legacy_laravel_redirect.php';
app_redirect_legacy_to_laravel('cxp/modules/accion4');

require dirname(__DIR__, 2) . '/includes/bootstrap.php';

use App\Shared\Application\ExternalScriptModuleController;
use App\Shared\Infrastructure\ExternalCommandRunner;
use App\Shared\Support\PhpView;

$controller = new ExternalScriptModuleController(
    new ExternalCommandRunner(),
    [
        'workspace_slug' => 'cxp',
        'module_slug' => 'cxp_accion4',
        'action_key' => 'accion4',
        'upload_field' => 'txt_file',
        'accepted_extensions' => ['txt'],
        'default_upload_name' => 'documento.txt',
        'fallback_base_name' => 'documento',
        'script_relative_path' => 'run_bot_accion4.js',
        'output_suffix' => '_accion4.xlsx',
        'output_flag' => '--output',
        'ignore_console_prefixes' => [
            'Auditoria JSON:',
            'Archivo leido:',
            'Excel generado (una sola hoja):',
        ],
        'generated_path_prefix' => 'Excel generado (una sola hoja):',
        'messages' => [
            'missing_file' => 'No se recibio el archivo TXT.',
            'upload_error' => 'Error al subir el archivo TXT.',
            'invalid_extension' => 'Solo se permiten archivos TXT.',
            'save_failure' => 'No se pudo guardar el TXT subido.',
            'script_missing' => 'No existe run_bot_accion4.js en el proyecto.',
            'generated_missing' => 'No se encontro el Excel generado.',
            'process_failed' => 'El proceso fallo.',
        ],
    ]
);
$viewData = $controller->handle($_SERVER, $_FILES);

$latestActions = app_latest_action_exports();
$missingActions = array_values(array_filter(
    $latestActions,
    static fn(array $item): bool => $item['latest'] === null
));
$canExportAllActions = count($missingActions) === 0;

$pageConfig = [
    'top_title' => (string)$viewData['workspace']['title'],
    'top_note' => 'Modulo operativo para procesamiento contable en produccion',
    'back_url' => app_url(),
    'back_label' => 'Volver al portal',
    'hero_chip' => 'Flujo productivo activo',
    'hero_title' => (string)$viewData['currentModule']['title'],
    'hero_lead' => 'Flujo en produccion: carga de un solo TXT MAYOR GENERAL IVA y generacion de Excel MAYOR IVA respetando exactamente la estructura del ejemplo manual en hoja unica.',
    'hero_note' => 'Salida de usuario: descarga directa de archivo .xlsx.',
    'hero_card_label' => 'Alcance actual',
    'hero_card_value' => 'Proceso principal de Cuentas por Pagar en operacion diaria.',
    'show_logo' => true,
    'sidebar_panels' => [
        [
            'title' => 'Mapa del proceso',
            'items' => [
                ['tag' => 'Entrada', 'title' => 'Un TXT MAYOR GENERAL IVA', 'description' => 'Documento fuente tipo MAYOR GENERAL IVA para devolver un solo Excel igual al ejemplo manual.'],
                ['tag' => 'Transformacion', 'title' => 'Replica exacta de estructura manual', 'description' => 'El sistema omite, reordena e inserta filas de formula para igualar el formato operativo esperado.'],
                ['tag' => 'Salida', 'title' => 'Excel final MAYOR IVA', 'description' => 'Archivo listo para descarga y uso operativo del departamento.'],
                ['tag' => 'Control interno', 'title' => 'Auditoria tecnica en storage', 'description' => 'Registro interno del proceso para revision del equipo tecnico.'],
                ['tag' => 'Siguiente accion', 'title' => 'Clasificacion TXT a Excel', 'description' => 'Proceso principal de proveedores (accion 1).', 'url' => app_url('modules/cxp_pdf/index.php'), 'link_label' => 'Abrir Accion 1'],
            ],
        ],
    ],
    'upload_title' => 'Procesar TXT MAYOR IVA',
    'upload_description' => 'Sube un solo TXT y el sistema devolvera un Excel MAYOR IVA igual al ejemplo manual cargado en el proyecto.',
    'form_label' => 'Archivo TXT MAYOR IVA',
    'form_field_name' => 'txt_file',
    'form_accept' => '.txt,text/plain',
    'button_label' => 'Generar Excel',
    'download_label' => 'Descargar Excel',
    'success_message' => $viewData['result'] !== null ? 'Archivo generado correctamente: ' . (string)$viewData['result']['excel_name'] : '',
    'stats_items' => [
        ['tag' => 'Salida publica', 'title' => 'Solo .xlsx', 'description' => 'La interfaz publica no expone el JSON de auditoria.'],
        ['tag' => 'Ubicacion de archivos', 'title' => 'storage/outputs', 'description' => 'Centralizado para respaldo y control del equipo.'],
        ['tag' => 'Disponibilidad', 'title' => 'Operacion interna', 'description' => 'Disenado para uso local del departamento contable.'],
        [
            'tag' => 'Consolidado total',
            'title' => $canExportAllActions ? 'Listo para descargar' : 'Pendiente de completar',
            'description' => $canExportAllActions
                ? 'Genera un solo Excel con el ultimo archivo producido en cada accion.'
                : 'Faltan salidas recientes para: ' . implode(', ', array_map(
                    static fn(array $item): string => (string)$item['label'],
                    $missingActions
                )) . '.',
            'url' => $canExportAllActions ? app_url('export_all_actions.php') : '',
            'button_label' => $canExportAllActions ? 'Descargar todas las acciones' : '',
        ],
    ],
    'secondary_panels' => [
        [
            'title' => 'Ultimo Excel por accion',
            'description' => 'Base usada para armar el consolidado general.',
            'items' => array_map(
                static function (array $action): array {
                    $latest = $action['latest'];
                    return [
                        'title' => (string)$action['label'],
                        'meta' => $latest !== null
                            ? (string)$latest['name'] . ' | ' . (string)$latest['time']
                            : 'Todavia no hay Excel generado para esta accion.',
                        'url' => $latest !== null
                            ? app_output_download_url((string)$latest['name'])
                            : (string)$action['module_url'],
                        'button_label' => $latest !== null ? 'Descargar' : 'Abrir accion',
                    ];
                },
                $latestActions
            ),
        ],
    ],
    'history_title' => 'Historial reciente',
    'history_description' => 'Archivos generados por este modulo de produccion. Se conservan solo los ultimos ' . app_output_retention_limit() . ' informes.',
    'history_empty_title' => 'Sin archivos generados todavia',
    'history_empty_description' => 'El historial se llena con cada ejecucion del proceso.',
];

(new PhpView())->render(
    app_join_path(app_root(), 'templates', 'cxp', 'script_module_page.php'),
    array_merge($viewData, ['pageConfig' => $pageConfig])
);
