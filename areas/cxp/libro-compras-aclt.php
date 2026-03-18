<?php
declare(strict_types=1);

require dirname(__DIR__, 2) . '/includes/bootstrap.php';

use App\Cxp\Application\CxpWindowViewModelFactory;
use App\Shared\Support\PhpView;

$viewData = (new CxpWindowViewModelFactory())->build('cxp', 'libro_compras_aclt', [
    'hero_card_label' => 'Area responsable',
    'hero_card_source' => 'workspace_title',
    'control_title' => 'Control Operativo',
    'control_description' => 'Esta ventana concentra los procesos del libro de compras y el consolidado general del flujo actual.',
    'process_title' => 'Procesos Disponibles',
    'process_description' => 'Las acciones quedan renombradas por su funcion real, sin alterar su funcionamiento interno.',
    'bundle_button_label' => 'Generar consolidado',
]);

(new PhpView())->render(
    app_join_path(app_root(), 'templates', 'cxp', 'windows', 'standard.php'),
    $viewData
);
