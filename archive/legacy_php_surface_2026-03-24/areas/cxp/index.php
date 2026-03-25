<?php
declare(strict_types=1);

require dirname(__DIR__, 2) . '/includes/legacy_laravel_redirect.php';
app_redirect_legacy_to_laravel('cxp');

require dirname(__DIR__, 2) . '/includes/bootstrap.php';

use App\Cxp\Application\CxpWorkspaceHomeViewModelFactory;
use App\Shared\Support\PhpView;

$viewData = (new CxpWorkspaceHomeViewModelFactory())->build('cxp', [
    'top_note' => 'Area responsable en produccion',
    'hero_chip' => 'Proceso responsable',
    'hero_card_label' => 'Acceso LAN',
    'hero_card_meta' => 'Ventanas activas: %d',
    'control_title' => 'Informacion Operativa',
    'control_description' => 'Esta area agrupa las ventanas de trabajo del responsable y conserva la salida publica controlada del proceso.',
    'windows_title' => 'Ventanas Disponibles',
    'windows_description' => 'Actualmente esta area opera con tres ventanas activas y queda lista para crecer sin mezclar responsabilidades entre procesos.',
]);

(new PhpView())->render(
    app_join_path(app_root(), 'templates', 'cxp', 'workspace_home.php'),
    $viewData
);
