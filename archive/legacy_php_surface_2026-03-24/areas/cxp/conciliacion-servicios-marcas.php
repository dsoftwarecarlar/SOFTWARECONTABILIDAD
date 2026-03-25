<?php
declare(strict_types=1);

require dirname(__DIR__, 2) . '/includes/legacy_laravel_redirect.php';
app_redirect_legacy_to_laravel('cxp/windows/conciliacion-servicios-marcas');

require dirname(__DIR__, 2) . '/includes/bootstrap.php';

use App\Cxp\ServiciosMarcas\Application\ServiciosMarcasWindowViewModelFactory;
use App\Shared\Support\PhpView;

$config = require app_join_path(app_root(), 'config', 'cxp', 'servicios_marcas.php');
$viewData = (new ServiciosMarcasWindowViewModelFactory())->build($config);

(new PhpView())->render(
    app_join_path(app_root(), 'templates', 'cxp', 'windows', 'servicios_marcas.php'),
    $viewData
);
