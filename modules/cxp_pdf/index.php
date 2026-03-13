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
        'module_slug' => 'cxp_pdf',
        'action_key' => 'accion1',
        'upload_field' => 'pdf_file',
        'accepted_extensions' => ['pdf'],
        'default_upload_name' => 'documento.pdf',
        'fallback_base_name' => 'documento',
        'script_relative_path' => 'run_bot.js',
        'output_suffix' => '_resultado.xlsx',
        'ignore_console_prefixes' => [
            'Auditoria JSON:',
            'PDF leido:',
            'Excel generado (una sola hoja):',
        ],
        'generated_path_prefix' => 'Excel generado (una sola hoja):',
        'messages' => [
            'missing_file' => 'No se recibio el archivo PDF.',
            'upload_error' => 'Error al subir el archivo PDF.',
            'invalid_extension' => 'Solo se permiten archivos PDF.',
            'save_failure' => 'No se pudo guardar el PDF subido.',
            'script_missing' => 'No existe run_bot.js en el proyecto.',
            'generated_missing' => 'No se encontro el Excel generado.',
            'process_failed' => 'El proceso fallo.',
        ],
    ]
);
$viewData = $controller->handle($_SERVER, $_FILES);

$pageConfig = [
    'top_title' => (string)$viewData['workspace']['title'],
    'top_note' => 'Modulo operativo para procesamiento contable en produccion',
    'back_url' => app_url(),
    'back_label' => 'Volver al portal',
    'hero_chip' => 'Flujo productivo activo',
    'hero_title' => (string)$viewData['currentModule']['title'],
    'hero_lead' => 'Flujo en produccion: carga de PDF del proveedor, extraccion de datos, clasificacion por tipo de movimiento y generacion de Excel final en hoja unica para uso del equipo contable.',
    'hero_note' => 'Salida de usuario: descarga directa de archivo .xlsx.',
    'hero_card_label' => 'Alcance actual',
    'hero_card_value' => 'Proceso principal de Cuentas por Pagar en operacion diaria.',
    'show_logo' => true,
    'sidebar_panels' => [
        [
            'title' => 'Mapa del proceso',
            'items' => [
                ['tag' => 'Entrada', 'title' => 'Archivo PDF del proveedor', 'description' => 'Documento fuente subido por el usuario del area.'],
                ['tag' => 'Transformacion', 'title' => 'Extraccion y clasificacion contable', 'description' => 'Validaciones de formato, ordenamiento y armado del libro en hoja unica.'],
                ['tag' => 'Salida', 'title' => 'Excel final LIBRO COMPRAS', 'description' => 'Archivo listo para descarga y uso operativo del departamento.'],
                ['tag' => 'Control interno', 'title' => 'Auditoria tecnica en storage', 'description' => 'Registro interno del proceso para revision del equipo tecnico.'],
                ['tag' => 'Siguiente accion', 'title' => 'TXT a Excel Retenciones', 'description' => 'Carga de TXT desordenado y conversion al formato exacto de ACCION2.', 'url' => app_url('modules/cxp_txt/index.php'), 'link_label' => 'Abrir Accion 2'],
            ],
        ],
    ],
    'upload_title' => 'Procesar archivo PDF',
    'upload_description' => 'Sube el documento y el sistema devolvera el Excel final de produccion para este flujo.',
    'form_label' => 'Archivo PDF de proveedores',
    'form_field_name' => 'pdf_file',
    'form_accept' => 'application/pdf',
    'button_label' => 'Generar Excel',
    'download_label' => 'Descargar Excel',
    'success_message' => $viewData['result'] !== null ? 'Archivo generado correctamente: ' . (string)$viewData['result']['excel_name'] : '',
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
