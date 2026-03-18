<?php
declare(strict_types=1);

/** @var array<string, mixed> $brand */
/** @var array<string, mixed> $workspace */
/** @var array<string, mixed> $window */
/** @var int $processCount */
/** @var array<string, mixed> $pageConfig */
/** @var array<int, string> $stylesheets */

$topNote = (string)($pageConfig['top_note'] ?? $workspace['title'] ?? '');
$backLabel = (string)($pageConfig['back_label'] ?? 'Volver al responsable');
$heroChip = (string)($pageConfig['hero_chip'] ?? 'Ventana operativa');
$heroNotePrefix = (string)($pageConfig['hero_note_prefix'] ?? 'Procesos disponibles');
$heroCardLabel = (string)($pageConfig['hero_card_label'] ?? 'Acceso LAN');
$heroCardSource = (string)($pageConfig['hero_card_source'] ?? 'public_root_url');
$heroCardValue = (string)($pageConfig['hero_card_value'] ?? ($heroCardSource === 'workspace_title'
    ? ($workspace['title'] ?? '')
    : app_public_root_url()));
$controlTitle = (string)($pageConfig['control_title'] ?? 'Control Operativo');
$controlDescription = (string)($pageConfig['control_description'] ?? '');
$processTitle = (string)($pageConfig['process_title'] ?? (!empty($window['bundle']) ? 'Procesos Disponibles' : 'Proceso Disponible'));
$processDescription = (string)($pageConfig['process_description'] ?? '');
$routeNote = (string)($window['route_note'] ?? ($pageConfig['route_fallback'] ?? 'Ruta de salida: storage/outputs.'));
$moduleButtonLabel = (string)($pageConfig['module_button_label'] ?? 'Abrir proceso');
$bundleButtonLabel = (string)($pageConfig['bundle_button_label'] ?? 'Generar consolidado');
?>
<!doctype html>
<html lang="es">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title><?= htmlspecialchars((string)$window['title']) ?></title>
    <?php foreach ($stylesheets as $stylesheet): ?>
        <link rel="stylesheet" href="<?= htmlspecialchars($stylesheet) ?>">
    <?php endforeach; ?>
</head>
<body class="app-page cxp-window-page">
<div class="page">
    <div class="topbar">
        <div>
            <span class="chip"><?= htmlspecialchars((string)$brand['platform']) ?></span>
            <strong><?= htmlspecialchars((string)$window['title']) ?></strong>
            <span class="note"><?= htmlspecialchars($topNote) ?></span>
        </div>
        <a class="back-link" href="<?= htmlspecialchars((string)($workspace['home_url'] ?? app_url())) ?>"><?= htmlspecialchars($backLabel) ?></a>
    </div>

    <section class="hero">
        <div class="hero-grid">
            <div>
                <span class="chip"><?= htmlspecialchars($heroChip) ?></span>
                <h1><?= htmlspecialchars((string)$window['title']) ?></h1>
                <p class="lead"><?= htmlspecialchars((string)$window['summary']) ?></p>
                <div class="hero-note"><?= htmlspecialchars($heroNotePrefix) ?>: <?= $processCount ?></div>
            </div>
            <div class="hero-card">
                <span><?= htmlspecialchars($heroCardLabel) ?></span>
                <strong><?= htmlspecialchars($heroCardValue) ?></strong>
            </div>
        </div>
    </section>

    <div class="layout">
        <aside class="panel">
            <h2><?= htmlspecialchars($controlTitle) ?></h2>
            <p><?= htmlspecialchars($controlDescription) ?></p>
            <div class="trace"><?= htmlspecialchars($routeNote) ?></div>
        </aside>

        <main class="panel">
            <h2><?= htmlspecialchars($processTitle) ?></h2>
            <p><?= htmlspecialchars($processDescription) ?></p>
            <div class="process-grid">
                <?php foreach (($window['modules'] ?? []) as $module): ?>
                    <article class="process-card">
                        <span><?= htmlspecialchars((string)$module['status']) ?></span>
                        <strong><?= htmlspecialchars((string)$module['title']) ?></strong>
                        <p><?= htmlspecialchars((string)$module['description']) ?></p>
                        <a class="cta" href="<?= htmlspecialchars((string)$module['url']) ?>"><?= htmlspecialchars($moduleButtonLabel) ?></a>
                    </article>
                <?php endforeach; ?>
                <?php if (!empty($window['bundle'])): ?>
                    <article class="process-card">
                        <span><?= htmlspecialchars((string)$window['bundle']['status']) ?></span>
                        <strong><?= htmlspecialchars((string)$window['bundle']['title']) ?></strong>
                        <p><?= htmlspecialchars((string)$window['bundle']['description']) ?></p>
                        <a class="cta" href="<?= htmlspecialchars((string)$window['bundle']['url']) ?>"><?= htmlspecialchars($bundleButtonLabel) ?></a>
                    </article>
                <?php endif; ?>
            </div>
        </main>
    </div>
</div>
</body>
</html>
