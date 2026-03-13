<?php
declare(strict_types=1);

require dirname(__DIR__, 2) . '/includes/bootstrap.php';

use App\Cxp\RepuestosTytserv\Application\RepuestosTytservModuleController;
use App\Cxp\RepuestosTytserv\Infrastructure\RepuestosTytservScriptGateway;
use App\Shared\Infrastructure\ExternalCommandRunner;
use App\Shared\Support\PhpView;

$config = require app_join_path(app_root(), 'config', 'cxp', 'repuestos_tytserv.php');
$controller = new RepuestosTytservModuleController(
    new RepuestosTytservScriptGateway(
        new ExternalCommandRunner(),
        (string)($config['paths']['script_path'] ?? '')
    ),
    $config
);

$viewData = $controller->handle($_SERVER, $_FILES);

(new PhpView())->render(
    app_join_path(app_root(), 'templates', 'cxp', 'repuestos_tytserv', 'index.php'),
    $viewData
);
