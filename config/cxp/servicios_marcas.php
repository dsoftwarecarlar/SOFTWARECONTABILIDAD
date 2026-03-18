<?php
declare(strict_types=1);

return [
    'paths' => [
        'template_dir' => app_first_existing_path(
            app_join_path(app_root(), 'resources', 'cxp', 'servicios_marcas', 'templates'),
            app_join_path(app_root(), 'outputs', 'EJEMPLOAMANOTAREA2')
        ),
    ],
    'module' => [
        'history_limit' => 4,
        'history_scan_limit' => 20,
        'accepted_extensions' => ['xls', 'xlsx'],
        'module_note' => 'Modulo operativo para conciliar servicios por marca',
        'hero' => [
            'chip' => 'Flujo productivo activo',
            'lead' => 'Sube los 4 archivos mensuales (2 TXT + 2 Excel). El sistema separa por marca y genera una plantilla final descargable.',
            'note' => 'Salida de usuario: un .xls por marca con sus hojas operativas.',
        ],
        'template_panel' => [
            'title' => 'Plantillas',
            'description' => 'Este modulo solo procesa servicios por marca y no modifica otras vistas. La salida genera 1 archivo por marca.',
            'base_tag' => 'Base mensual',
            'base_title' => 'Plantillas del mes',
            'base_description' => 'La carpeta mensual contiene las plantillas base por marca.',
        ],
        'upload_panel' => [
            'title' => 'Cargar archivos',
            'description' => 'Sube REP FACTURACION (txt), NOTA DE CREDITO (txt), PX (xlsx) y REP VENTAS (xls/xlsx).',
            'button_label' => 'Procesar y generar plantillas',
            'processing_label' => 'Procesando en segundo plano...',
            'retention_tag' => 'Retencion',
            'retention_title' => '1 archivo por marca',
            'retention_description' => 'La limpieza automatica conserva solo salidas recientes en `storage/outputs`.',
            'pending_hint' => 'Puedes dejar esta pagina abierta. Se actualizara sola cuando termine.',
        ],
        'history_panel' => [
            'title' => 'Historial Reciente',
            'description' => 'Se muestran solo 4 salidas: el archivo mas reciente disponible de cada marca.',
            'empty_title' => 'No hay archivos generados aun.',
            'empty_description' => 'La primera ejecucion aparecera aqui.',
        ],
        'stop' => [
            'button_label' => 'Parar todos los procesos',
            'confirm_message' => 'Se solicitara detener todos los procesos activos de esta accion.',
            'processing_label' => 'Solicitando parada...',
            'meta_template' => 'Procesos activos: %d - Historial visible: ultimo archivo por marca.',
        ],
        'poll' => [
            'interval_ms' => 4000,
            'initial_delay_ms' => 2000,
            'reconnecting_message' => 'Estado: reconectando con el proceso...',
            'status_error' => 'No se pudo consultar el estado del proceso.',
        ],
    ],
    'window' => [
        'hero_chip' => 'Ventana operativa',
        'control_title' => 'Control Operativo',
        'control_description' => 'Esta ventana concentra el proceso mensual que separa el reporte de servicios por marca y genera una plantilla final descargable para cada empresa.',
        'process_title' => 'Proceso Disponible',
        'process_description' => 'El flujo queda separado como segunda ventana operativa, sin mezclarlo con el libro de compras ACLT.',
        'route_fallback' => 'Ruta de salida: storage/outputs.',
    ],
    'brands' => [
        ['key' => 'changan', 'label' => 'CHANGAN', 'description' => 'Datos filtrados por marca en su libro operativo.'],
        ['key' => 'peug', 'label' => 'PEUGEOT', 'description' => 'Facturas y notas separadas en la base visible del mes.'],
        ['key' => 'szk', 'label' => 'SUZUKI', 'description' => 'Consolida SUZUKI AMBATO y SUZUKI RIOBAMBA.'],
        ['key' => 'tyt', 'label' => 'MATRIZ', 'description' => 'Respeta el formato mensual definido para MATRIZ.'],
    ],
];
