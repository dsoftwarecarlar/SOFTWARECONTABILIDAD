<?php
declare(strict_types=1);

require dirname(__DIR__, 2) . '/includes/bootstrap.php';
require dirname(__DIR__, 2) . '/includes/servicios_marcas_job_runner.php';

use App\Cxp\ServiciosMarcas\Application\ServiciosMarcasModuleController;
use App\Cxp\ServiciosMarcas\Domain\HistoryLabelResolver;
use App\Cxp\ServiciosMarcas\Infrastructure\ServiciosMarcasJobGateway;
use App\Shared\Support\PhpView;

$config = require app_join_path(app_root(), 'config', 'cxp', 'servicios_marcas.php');
$jobsDir = app_ensure_dir(app_storage_path('jobs'));

$controller = new ServiciosMarcasModuleController(
    new ServiciosMarcasJobGateway($jobsDir),
    new HistoryLabelResolver(servicios_job_output_config()),
    $config
);

$viewData = $controller->handle($_SERVER, $_GET, $_POST, $_FILES);

(new PhpView())->render(
    app_join_path(app_root(), 'templates', 'cxp', 'servicios_marcas', 'index.php'),
    $viewData
);
