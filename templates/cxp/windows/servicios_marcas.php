<?php
declare(strict_types=1);

/** @var array<string, mixed> $brand */
/** @var array<string, mixed> $workspace */
/** @var array<string, mixed> $window */
/** @var int $processCount */
/** @var array<string, mixed> $pageConfig */
/** @var array<int, string> $stylesheets */
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
<body class="app-page servicios-marcas-window-page">
<div class="page">
    <div class="topbar">
        <div>
            <span class="chip"><?= htmlspecialchars((string)$brand['platform']) ?></span>
            <strong><?= htmlspecialchars((string)$window['title']) ?></strong>
            <span class="note"><?= htmlspecialchars((string)$workspace['title']) ?></span>
        </div>
        <a class="back-link" href="<?= htmlspecialchars((string)($workspace['home_url'] ?? app_url())) ?>">Volver al responsable</a>
    </div>

    <section class="hero">
        <div class="hero-grid">
            <div>
                <span class="chip"><?= htmlspecialchars((string)($pageConfig['hero_chip'] ?? 'Ventana operativa')) ?></span>
                <h1><?= htmlspecialchars((string)$window['title']) ?></h1>
                <p class="lead"><?= htmlspecialchars((string)$window['summary']) ?></p>
                <div class="hero-note">Procesos disponibles: <?= $processCount ?></div>
            </div>
            <div class="hero-card">
                <span>Acceso LAN</span>
                <strong><?= htmlspecialchars(app_public_root_url()) ?></strong>
            </div>
        </div>
    </section>

    <div class="layout">
        <aside class="panel">
            <h2><?= htmlspecialchars((string)($pageConfig['control_title'] ?? 'Control Operativo')) ?></h2>
            <p><?= htmlspecialchars((string)($pageConfig['control_description'] ?? '')) ?></p>
            <div class="trace"><?= htmlspecialchars((string)($window['route_note'] ?? ($pageConfig['route_fallback'] ?? 'Ruta de salida: storage/outputs.'))) ?></div>
        </aside>

        <main class="panel">
            <h2><?= htmlspecialchars((string)($pageConfig['process_title'] ?? 'Proceso Disponible')) ?></h2>
            <p><?= htmlspecialchars((string)($pageConfig['process_description'] ?? '')) ?></p>
            <div class="process-grid">
                <?php foreach (($window['modules'] ?? []) as $module): ?>
                    <article class="process-card">
                        <span><?= htmlspecialchars((string)$module['status']) ?></span>
                        <strong><?= htmlspecialchars((string)$module['title']) ?></strong>
                        <p><?= htmlspecialchars((string)$module['description']) ?></p>
                        <a class="cta" href="<?= htmlspecialchars((string)$module['url']) ?>">Abrir proceso</a>
                    </article>
                <?php endforeach; ?>
            </div>
        </main>
    </div>
</div>
</body>
</html>
