<?php
declare(strict_types=1);

require dirname(__DIR__, 2) . '/includes/legacy_laravel_redirect.php';
app_redirect_legacy_to_laravel('cxp/windows/facturacion-repuestos-tytserv');

require dirname(__DIR__, 2) . '/includes/bootstrap.php';

use App\Cxp\Application\CxpWindowViewModelFactory;
use App\Shared\Support\PhpView;

$viewData = (new CxpWindowViewModelFactory())->build('cxp', 'facturacion_repuestos_tytserv', [
    'control_title' => 'Control Operativo',
    'control_description' => 'Esta ventana procesa los 4 archivos de ventas por marca y los unifica en una sola plantilla mensual.',
    'process_title' => 'Proceso Disponible',
    'process_description' => 'El flujo queda aislado de las otras vistas para no afectar Libro de Compras ACLT ni conciliacion de servicios.',
]);

(new PhpView())->render(
    app_join_path(app_root(), 'templates', 'cxp', 'windows', 'standard.php'),
    $viewData
);
