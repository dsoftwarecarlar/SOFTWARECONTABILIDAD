<?php
declare(strict_types=1);

return [
    'paths' => [
        'template_dir' => app_first_existing_path(
            app_join_path(app_root(), 'resources', 'cxp', 'repuestos_tytserv', 'templates'),
            app_join_path(app_root(), 'outputs', 'EJEMPLOAMANOTAREA3')
        ),
        // Ruta productiva actual del modulo web.
        // run_repuestos_tytserv.ps1 queda como fallback/manual probe verificado aparte.
        'script_path' => app_join_path(app_root(), 'scripts', 'cxp', 'repuestos_tytserv', 'process.js'),
    ],
    'module' => [
        'window_slug' => 'facturacion_repuestos_tytserv',
        'template_file' => 'FACTURACION REPUESTOS TYTSERV FEBRERO 2026.xlsx',
        'accepted_extensions' => ['xls', 'xlsx'],
        'file_fields' => [
            [
                'field' => 'excel_tyt',
                'label' => 'MATRIZ (RepLibroVentasGeneral)',
                'script_flag' => '-InputTyt',
                'summary_key' => 'tyt',
                'summary_label' => 'MATRIZ',
            ],
            [
                'field' => 'excel_peug',
                'label' => 'PEUGEOT (RepLibroVentasGeneral)',
                'script_flag' => '-InputPeug',
                'summary_key' => 'peug',
                'summary_label' => 'PEUGEOT',
            ],
            [
                'field' => 'excel_chgn',
                'label' => 'CHANGAN (RepLibroVentasGeneral)',
                'script_flag' => '-InputChgn',
                'summary_key' => 'chgn',
                'summary_label' => 'CHANGAN',
            ],
            [
                'field' => 'excel_szk',
                'label' => 'SUZUKI (RepLibroVentasGeneral)',
                'script_flag' => '-InputSzk',
                'summary_key' => 'szk',
                'summary_label' => 'SUZUKI',
            ],
        ],
        'page' => [
            'top_note' => 'Carga y procesamiento mensual por 4 marcas',
            'hero_chip' => 'Flujo mensual',
            'hero_lead' => 'Carga los 4 archivos RepLibroVentasGeneral del mes y genera la salida final en el formato de la plantilla manual.',
            'hero_note' => 'Salida: un archivo .xlsx listo para descarga.',
            'inputs_title' => 'Entradas requeridas',
            'retention_tag' => 'Retencion',
            'retention_description' => 'Se elimina automaticamente lo mas antiguo.',
            'process_title' => 'Procesar plantilla',
            'process_description' => 'Sube los 4 archivos del mes para generar el consolidado final.',
            'button_label' => 'Procesar y generar reporte',
            'output_tag' => 'Salida',
            'output_title' => 'storage/outputs',
            'output_description' => 'Prefijo de archivos: repuestos_tytserv_',
            'history_title' => 'Historial reciente',
            'history_empty_title' => 'No hay reportes generados aun.',
            'history_empty_description' => 'La primera ejecucion aparecera aqui.',
        ],
    ],
];
