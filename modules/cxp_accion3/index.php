<?php
declare(strict_types=1);

require dirname(__DIR__, 2) . '/includes/bootstrap.php';

use App\Shared\Application\ExternalScriptModuleController;
use App\Shared\Infrastructure\ExternalCommandRunner;
use App\Shared\Support\PhpView;

$controller = new ExternalScriptModuleController(
    new ExternalCommandRunner(),
    [
        'workspace_slug' => 'cxp',
        'module_slug' => 'cxp_accion3',
        'action_key' => 'accion3',
        'upload_field' => 'txt_file',
        'allow_multiple' => true,
        'accepted_extensions' => ['txt'],
        'default_upload_name' => 'documento.txt',
        'fallback_base_name' => 'documento',
        'script_relative_path' => 'run_bot_accion3.js',
        'output_suffix' => '_accion3.xlsx',
        'output_flag' => '--output',
        'ignore_console_prefixes' => [
            'Auditoria JSON:',
            'Archivo leido:',
            'Excel generado (una sola hoja):',
        ],
        'generated_path_prefix' => 'Excel generado (una sola hoja):',
        'output_base_resolver' => static function (array $savedInputs, string $timestamp): string {
            $baseName = (string)($savedInputs[0]['safe_base'] ?? 'documento');
            if (count($savedInputs) > 1) {
                $baseName .= '_lote_' . count($savedInputs);
            }

            return $baseName;
        },
        'result_builder' => static fn(array $savedInputs, array $execution, string $generatedName): array => [
            'source_count' => count($savedInputs),
        ],
        'messages' => [
            'missing_file' => 'No se recibieron archivos TXT.',
            'invalid_upload_shape' => 'Formato de carga invalido.',
            'invalid_batch_shape' => 'La carga multiple no llego correctamente.',
            'upload_error' => 'Error al subir uno de los archivos TXT.',
            'empty_selection' => 'Selecciona al menos un archivo TXT.',
            'invalid_extension' => 'Solo se permiten archivos TXT.',
            'save_failure' => 'No se pudo guardar uno de los TXT subidos.',
            'script_missing' => 'No existe run_bot_accion3.js en el proyecto.',
            'generated_missing' => 'No se encontro el Excel generado.',
            'process_failed' => 'El proceso fallo.',
        ],
    ]
);
$viewData = $controller->handle($_SERVER, $_FILES);

$successMessage = '';
if ($viewData['result'] !== null) {
    $successMessage = 'Archivo generado correctamente: ' . (string)$viewData['result']['excel_name'];
    if (!empty($viewData['result']['generated_at'])) {
        $successMessage .= ' | Generado: ' . (string)$viewData['result']['generated_at'];
    }
    if ((int)($viewData['result']['source_count'] ?? 0) > 1) {
        $successMessage .= ' | TXT consolidados: ' . (int)$viewData['result']['source_count'];
    }
    if (!empty($viewData['result']['output_origin_note'])) {
        $successMessage .= ' | ' . (string)$viewData['result']['output_origin_note'];
    }
}

$pageConfig = [
    'top_title' => (string)$viewData['workspace']['title'],
    'top_note' => 'Modulo operativo para procesamiento contable en produccion',
    'back_url' => app_url(),
    'back_label' => 'Volver al portal',
    'hero_chip' => 'Flujo productivo activo',
    'hero_title' => (string)$viewData['currentModule']['title'],
    'hero_lead' => 'Flujo en produccion: carga de TXT MAYOR GENERAL, extraccion de movimientos contables y generacion de Excel MAYOR RET manteniendo formato operativo en hoja unica.',
    'hero_note' => 'Salida de usuario: descarga directa de archivo .xlsx.',
    'hero_card_label' => 'Alcance actual',
    'hero_card_value' => 'Proceso principal de Cuentas por Pagar en operacion diaria.',
    'show_logo' => true,
    'sidebar_panels' => [
        [
            'title' => 'Mapa del proceso',
            'items' => [
                ['tag' => 'Entrada', 'title' => 'Uno o varios TXT MAYOR GENERAL', 'description' => 'Documentos fuente tipo CON_MAYORGEN2ACCION3 para consolidar en una sola salida.'],
                ['tag' => 'Transformacion', 'title' => 'Extraccion de movimientos y resumen', 'description' => 'Normalizacion de filas y armado de salida en formato MAYOR RET.'],
                ['tag' => 'Salida', 'title' => 'Excel final MAYOR RET', 'description' => 'Archivo listo para descarga y uso operativo del departamento.'],
                ['tag' => 'Control interno', 'title' => 'Auditoria tecnica en storage', 'description' => 'Registro interno del proceso para revision del equipo tecnico.'],
                ['tag' => 'Siguiente accion', 'title' => 'Clasificacion TXT a Excel', 'description' => 'Proceso principal de proveedores (accion 1).', 'url' => app_url('modules/cxp_pdf/index.php'), 'link_label' => 'Abrir Accion 1'],
            ],
        ],
    ],
    'upload_title' => 'Procesar TXT MAYOR GENERAL',
    'upload_description' => 'Sube uno o varios TXT y el sistema devolvera un solo Excel consolidado para esta accion.',
    'form_label' => 'Archivos TXT MAYOR GENERAL',
    'form_field_name' => 'txt_file[]',
    'form_accept' => '.txt,text/plain',
    'form_multiple' => true,
    'button_label' => 'Generar Excel',
    'download_label' => 'Descargar Excel',
    'success_message' => $successMessage,
    'stats_items' => [
        ['tag' => 'Salida publica', 'title' => 'Solo .xlsx', 'description' => 'La interfaz publica no expone el JSON de auditoria.'],
        ['tag' => 'Ubicacion de archivos', 'title' => 'storage/outputs', 'description' => 'Centralizado para respaldo y control del equipo.'],
        ['tag' => 'Disponibilidad', 'title' => 'Operacion interna', 'description' => 'Disenado para uso local del departamento contable.'],
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
