<?php
declare(strict_types=1);

/** @var array<string, mixed> $brand */
/** @var array<string, mixed> $workspace */
/** @var int $windowCount */
/** @var string $publicRootUrl */
/** @var array<string, mixed> $pageConfig */
/** @var array<int, string> $stylesheets */
?>
<!doctype html>
<html lang="es">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title><?= htmlspecialchars((string)$workspace['title']) ?></title>
    <?php foreach ($stylesheets as $stylesheet): ?>
        <link rel="stylesheet" href="<?= htmlspecialchars($stylesheet) ?>">
    <?php endforeach; ?>
</head>
<body class="app-page cxp-workspace-page">
<div class="page">
    <div class="topbar">
        <div>
            <span class="chip"><?= htmlspecialchars((string)$brand['platform']) ?></span>
            <strong><?= htmlspecialchars((string)$workspace['title']) ?></strong>
            <span class="note"><?= htmlspecialchars((string)($pageConfig['top_note'] ?? 'Area responsable en produccion')) ?></span>
        </div>
        <a class="back-link" href="<?= htmlspecialchars(app_url()) ?>"><?= htmlspecialchars((string)($pageConfig['back_label'] ?? 'Volver al portal')) ?></a>
    </div>

    <section class="hero">
        <div class="hero-grid">
            <div>
                <span class="chip"><?= htmlspecialchars((string)($pageConfig['hero_chip'] ?? 'Proceso responsable')) ?></span>
                <h1><?= htmlspecialchars((string)$workspace['title']) ?></h1>
                <p class="lead"><?= htmlspecialchars((string)$workspace['summary']) ?></p>
                <div class="hero-note"><?= htmlspecialchars((string)$workspace['person']) ?></div>
            </div>
            <div class="hero-card">
                <span><?= htmlspecialchars((string)($pageConfig['hero_card_label'] ?? 'Acceso LAN')) ?></span>
                <strong><?= htmlspecialchars($publicRootUrl) ?></strong>
                <div class="meta"><?= htmlspecialchars(sprintf((string)($pageConfig['hero_card_meta'] ?? 'Ventanas activas: %d'), $windowCount)) ?></div>
            </div>
        </div>
    </section>

    <div class="layout">
        <section class="panel">
            <h2><?= htmlspecialchars((string)($pageConfig['control_title'] ?? 'Informacion Operativa')) ?></h2>
            <p><?= htmlspecialchars((string)($pageConfig['control_description'] ?? '')) ?></p>
            <div class="trace">
                <strong>Ruta de salida:</strong> `storage/outputs`.<br>
                <strong>Descarga publica:</strong> solo archivo `.xlsx`.<br>
                <strong>Auditoria:</strong> archivo interno para control del area.<br>
                <strong>Ingreso directo:</strong> <a href="<?= htmlspecialchars($publicRootUrl) ?>">Ingresar</a>
            </div>
        </section>

        <section class="panel">
            <h2><?= htmlspecialchars((string)($pageConfig['windows_title'] ?? 'Ventanas Disponibles')) ?></h2>
            <p><?= htmlspecialchars((string)($pageConfig['windows_description'] ?? '')) ?></p>
            <div class="window-list">
                <?php foreach (($workspace['windows'] ?? []) as $window): ?>
                    <article class="window-card">
                        <span>Ventana</span>
                        <strong><?= htmlspecialchars((string)$window['title']) ?></strong>
                        <p><?= htmlspecialchars((string)$window['summary']) ?></p>
                        <a class="cta" href="<?= htmlspecialchars((string)($window['url'] ?? $workspace['home_url'])) ?>">Ingresar</a>
                    </article>
                <?php endforeach; ?>
            </div>
        </section>
    </div>
</div>
</body>
</html>
