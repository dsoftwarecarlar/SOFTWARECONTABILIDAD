<?php

declare(strict_types=1);

return [
    'root' => env('PYTHON_SERVICES_ROOT', dirname(__DIR__, 2) . DIRECTORY_SEPARATOR . 'python_services'),
    'binary' => env('PYTHON_BINARY', 'python'),
    'contract' => [
        'request_format' => 'json-manifest',
        'stdout_success_prefix' => 'OUTPUT|',
        'stdout_error_prefix' => 'ERROR|',
    ],
    'planned_processors' => [
        'cxp_actions' => [
            'entrypoint' => 'python_services/processors/cxp_actions',
            'scope' => 'Accion 1-4 y consolidado',
        ],
        'repuestos_tytserv' => [
            'entrypoint' => 'python_services/processors/repuestos_tytserv/process.py',
            'scope' => 'Piloto futuro cuando Node deje de ser necesario como runtime primario.',
        ],
        'servicios_marcas' => [
            'entrypoint' => 'python_services/processors/servicios_marcas/dispatch.py',
            'scope' => 'Orquestacion del job desde Laravel hacia el runner heredado PowerShell / Excel COM.',
        ],
    ],
];
