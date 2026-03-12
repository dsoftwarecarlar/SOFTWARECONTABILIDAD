<?php
declare(strict_types=1);

require dirname(__DIR__, 2) . '/includes/app.php';

$brand = app_brand();
$workspace = app_workspace_by_slug('cxp');
if ($workspace === null) {
    throw new RuntimeException('No se encontro la configuracion del area.');
}

$publicRootUrl = app_public_root_url();
$windowCount = count($workspace['windows'] ?? []);
?>
<!doctype html>
<html lang="es">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title><?= htmlspecialchars($workspace['title']) ?></title>
    <style>
        :root {
            --ink: #18242c;
            --muted: #5f6970;
            --line: rgba(24, 36, 44, 0.11);
            --panel: rgba(255, 252, 246, 0.92);
            --panel-strong: rgba(255, 255, 255, 0.96);
            --deep: #112530;
            --accent: #106f66;
            --accent-2: #bf7534;
            --bg-1: #ece3d3;
            --bg-2: #e5eded;
            --shadow: 0 20px 48px rgba(17, 28, 37, 0.10);
        }
        * { box-sizing: border-box; }
        body {
            margin: 0;
            color: var(--ink);
            background:
                radial-gradient(circle at 9% 11%, rgba(191, 117, 52, 0.22), transparent 23%),
                radial-gradient(circle at 86% 12%, rgba(16, 111, 102, 0.18), transparent 24%),
                linear-gradient(180deg, var(--bg-1) 0%, var(--bg-2) 100%);
            font-family: "Trebuchet MS", "Lucida Sans Unicode", sans-serif;
        }
        .page {
            max-width: 1240px;
            margin: 0 auto;
            padding: 28px 18px 44px;
        }
        .topbar {
            display: flex;
            justify-content: space-between;
            align-items: center;
            gap: 14px;
            margin-bottom: 16px;
        }
        .chip {
            display: inline-flex;
            align-items: center;
            padding: 6px 10px;
            border-radius: 999px;
            background: rgba(16, 111, 102, 0.10);
            color: var(--accent);
            font-size: 12px;
            font-weight: 700;
            letter-spacing: 0.08em;
            text-transform: uppercase;
        }
        .topbar strong {
            display: block;
            margin-top: 8px;
            font-family: "Franklin Gothic Medium", "Arial Narrow", sans-serif;
            font-size: 28px;
            letter-spacing: -0.03em;
        }
        .topbar span.note {
            display: block;
            margin-top: 3px;
            color: var(--muted);
            font-size: 14px;
        }
        .back-link {
            color: var(--deep);
            text-decoration: none;
            font-weight: 700;
        }
        .hero {
            position: relative;
            overflow: hidden;
            border-radius: 34px;
            border: 1px solid var(--line);
            box-shadow: var(--shadow);
            background: linear-gradient(135deg, rgba(255, 252, 246, 0.98), rgba(243, 235, 220, 0.84));
            padding: 30px;
        }
        .hero::after {
            content: "";
            position: absolute;
            right: -68px;
            top: -80px;
            width: 290px;
            height: 290px;
            border-radius: 50%;
            background: radial-gradient(circle, rgba(16, 111, 102, 0.16), transparent 68%);
        }
        .hero-grid {
            position: relative;
            z-index: 1;
            display: grid;
            grid-template-columns: minmax(0, 1fr) 320px;
            gap: 22px;
            align-items: center;
        }
        h1 {
            margin: 14px 0 12px;
            font-family: "Bodoni MT", "Book Antiqua", serif;
            font-size: clamp(38px, 5.8vw, 72px);
            line-height: 0.92;
            letter-spacing: -0.05em;
            max-width: 780px;
        }
        .lead {
            margin: 0;
            color: var(--muted);
            max-width: 760px;
            font-size: 17px;
            line-height: 1.62;
        }
        .hero-note {
            margin-top: 18px;
            display: inline-flex;
            align-items: center;
            padding: 10px 14px;
            border-radius: 16px;
            background: rgba(16, 111, 102, 0.09);
            color: var(--deep);
            font-weight: 700;
        }
        .hero-card {
            padding: 20px;
            border-radius: 24px;
            background: linear-gradient(180deg, rgba(21, 38, 48, 0.95), rgba(14, 28, 37, 0.96));
            color: #f7f2e9;
            box-shadow: inset 0 0 0 1px rgba(255, 255, 255, 0.08);
        }
        .hero-card span {
            display: block;
            color: rgba(247, 242, 233, 0.72);
            font-size: 12px;
            font-weight: 700;
            letter-spacing: 0.08em;
            text-transform: uppercase;
        }
        .hero-card strong {
            display: block;
            margin-top: 10px;
            font-size: 26px;
            line-height: 1.05;
        }
        .layout {
            margin-top: 26px;
            display: grid;
            grid-template-columns: 1.05fr 0.95fr;
            gap: 16px;
        }
        .panel {
            padding: 24px;
            border-radius: 28px;
            border: 1px solid var(--line);
            background: var(--panel-strong);
            box-shadow: var(--shadow);
        }
        .panel h2 {
            margin: 0 0 8px;
            font-family: "Franklin Gothic Medium", "Arial Narrow", sans-serif;
            font-size: 30px;
            letter-spacing: -0.03em;
        }
        .panel p {
            margin: 0;
            color: var(--muted);
            line-height: 1.58;
        }
        .trace {
            margin-top: 14px;
            padding: 14px 16px;
            border-radius: 18px;
            border: 1px dashed rgba(22, 35, 43, 0.18);
            background: rgba(255, 255, 255, 0.72);
            color: var(--muted);
            line-height: 1.55;
        }
        .trace strong {
            color: var(--deep);
        }
        .trace a {
            color: var(--accent);
            font-weight: 700;
            text-decoration: none;
        }
        .window-list {
            display: grid;
            gap: 14px;
            margin-top: 12px;
        }
        .window-card {
            padding: 18px;
            border-radius: 20px;
            border: 1px solid rgba(24, 36, 44, 0.12);
            background: rgba(248, 250, 249, 0.92);
        }
        .window-card span {
            display: inline-flex;
            padding: 5px 10px;
            border-radius: 999px;
            background: rgba(16, 111, 102, 0.09);
            color: var(--accent);
            font-size: 11px;
            font-weight: 700;
            letter-spacing: 0.08em;
            text-transform: uppercase;
        }
        .window-card strong {
            display: block;
            margin-top: 10px;
            font-size: 24px;
            line-height: 1.04;
        }
        .window-card p {
            margin-top: 8px;
        }
        .cta {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            margin-top: 14px;
            padding: 12px 18px;
            border-radius: 999px;
            background: linear-gradient(135deg, var(--accent) 0%, var(--deep) 100%);
            color: #fff;
            text-decoration: none;
            font-weight: 700;
        }
        @media (max-width: 980px) {
            .hero-grid,
            .layout {
                grid-template-columns: 1fr;
            }
        }
    </style>
</head>
<body>
<div class="page">
    <div class="topbar">
        <div>
            <span class="chip"><?= htmlspecialchars($brand['platform']) ?></span>
            <strong><?= htmlspecialchars($workspace['title']) ?></strong>
            <span class="note">Area responsable en produccion</span>
        </div>
        <a class="back-link" href="<?= htmlspecialchars(app_url()) ?>">Volver al portal</a>
    </div>

    <section class="hero">
        <div class="hero-grid">
            <div>
                <span class="chip">Proceso responsable</span>
                <h1><?= htmlspecialchars($workspace['title']) ?></h1>
                <p class="lead"><?= htmlspecialchars($workspace['summary']) ?></p>
                <div class="hero-note"><?= htmlspecialchars($workspace['person']) ?></div>
            </div>
            <div class="hero-card">
                <span>Acceso LAN</span>
                <strong><?= htmlspecialchars($publicRootUrl) ?></strong>
            </div>
        </div>
    </section>

    <div class="layout">
        <section class="panel">
            <h2>Informacion Operativa</h2>
            <p>Esta area agrupa las ventanas de trabajo del responsable y conserva la salida publica controlada del proceso.</p>
            <div class="trace">
                <strong>Ruta de salida:</strong> `storage/outputs`.<br>
                <strong>Descarga publica:</strong> solo archivo `.xlsx`.<br>
                <strong>Auditoria:</strong> archivo interno para control del area.<br>
                <strong>Ingreso directo:</strong> <a href="<?= htmlspecialchars($publicRootUrl) ?>">Ingresar</a>
            </div>
        </section>

        <section class="panel">
            <h2>Ventanas Disponibles</h2>
            <p>Actualmente esta area opera con una ventana principal y queda lista para crecer con nuevas ventanas sin tocar las acciones existentes.</p>
            <div class="window-list">
                <?php foreach (($workspace['windows'] ?? []) as $window): ?>
                    <article class="window-card">
                        <span>Ventana</span>
                        <strong><?= htmlspecialchars($window['title']) ?></strong>
                        <p><?= htmlspecialchars($window['summary']) ?></p>
                        <a class="cta" href="<?= htmlspecialchars((string)($window['url'] ?? $workspace['home_url'])) ?>">Ingresar</a>
                    </article>
                <?php endforeach; ?>
            </div>
        </section>
    </div>
</div>
</body>
</html>
