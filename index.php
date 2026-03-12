<?php
declare(strict_types=1);

require __DIR__ . '/includes/app.php';

$brand = app_brand();
$workspaces = app_workspaces();
$activeCount = 0;
$windowCount = 0;
$bundleCount = 0;
$publicRootUrl = app_public_root_url();

foreach ($workspaces as $workspace) {
    foreach (($workspace['modules'] ?? []) as $module) {
        if (($module['state'] ?? '') === 'active') {
            $activeCount += 1;
        }
    }

    $windowCount += count($workspace['windows'] ?? []);
    foreach (($workspace['windows'] ?? []) as $window) {
        if (($window['bundle']['state'] ?? '') === 'active') {
            $bundleCount += 1;
        }
    }
}
?>
<!doctype html>
<html lang="es">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title><?= htmlspecialchars($brand['platform']) ?></title>
    <style>
        :root {
            --ink: #16232b;
            --muted: #5e6a71;
            --line: rgba(22, 35, 43, 0.12);
            --panel: rgba(255, 252, 246, 0.92);
            --panel-strong: rgba(255, 255, 255, 0.96);
            --deep: #102431;
            --accent: #0f6f67;
            --accent-2: #be7433;
            --bg-1: #efe5d5;
            --bg-2: #e2ebed;
            --shadow: 0 20px 48px rgba(15, 27, 36, 0.10);
        }
        * { box-sizing: border-box; }
        body {
            margin: 0;
            color: var(--ink);
            background:
                radial-gradient(circle at 11% 14%, rgba(190, 116, 51, 0.22), transparent 24%),
                radial-gradient(circle at 84% 12%, rgba(15, 111, 103, 0.18), transparent 24%),
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
            background: rgba(15, 111, 103, 0.10);
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
            color: var(--muted);
            font-size: 14px;
            margin-top: 3px;
        }
        .status-pill {
            padding: 8px 12px;
            border-radius: 999px;
            border: 1px solid rgba(15, 111, 103, 0.22);
            background: rgba(255, 255, 255, 0.74);
            color: var(--deep);
            font-weight: 700;
            font-size: 12px;
            letter-spacing: 0.08em;
            text-transform: uppercase;
        }
        .hero {
            position: relative;
            overflow: hidden;
            border-radius: 34px;
            border: 1px solid var(--line);
            box-shadow: var(--shadow);
            background: linear-gradient(135deg, rgba(255, 252, 246, 0.98), rgba(245, 235, 218, 0.84));
            padding: 30px;
        }
        .hero::after {
            content: "";
            position: absolute;
            right: -72px;
            bottom: -86px;
            width: 300px;
            height: 300px;
            border-radius: 50%;
            background: radial-gradient(circle, rgba(15, 111, 103, 0.15), transparent 68%);
        }
        .hero-grid {
            position: relative;
            z-index: 1;
            display: grid;
            grid-template-columns: minmax(0, 1fr) 300px;
            gap: 24px;
            align-items: center;
        }
        h1 {
            margin: 14px 0 12px;
            font-family: "Bodoni MT", "Book Antiqua", serif;
            font-size: clamp(40px, 6vw, 72px);
            line-height: 0.92;
            letter-spacing: -0.05em;
            max-width: 760px;
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
        .logo-stage {
            padding: 18px;
            border-radius: 30px;
            background: linear-gradient(180deg, rgba(22, 38, 48, 0.95), rgba(14, 27, 36, 0.96));
            box-shadow: inset 0 0 0 1px rgba(255, 255, 255, 0.08);
            display: grid;
            place-items: center;
        }
        .logo-frame {
            width: 192px;
            height: 192px;
            padding: 18px;
            border-radius: 40px;
            display: grid;
            place-items: center;
            background: linear-gradient(180deg, rgba(255, 255, 255, 0.98), rgba(247, 238, 223, 0.95));
            box-shadow: 0 20px 36px rgba(0, 0, 0, 0.22);
        }
        .logo-frame img {
            width: 100%;
            height: 100%;
            object-fit: contain;
        }
        .stats {
            margin-top: 18px;
            display: grid;
            grid-template-columns: repeat(4, minmax(0, 1fr));
            gap: 14px;
        }
        .stat {
            padding: 18px;
            border-radius: 22px;
            border: 1px solid var(--line);
            background: var(--panel);
            box-shadow: var(--shadow);
        }
        .stat span {
            display: block;
            color: var(--muted);
            font-size: 12px;
            font-weight: 700;
            letter-spacing: 0.08em;
            text-transform: uppercase;
        }
        .stat strong {
            display: block;
            margin-top: 8px;
            font-family: "Franklin Gothic Medium", "Arial Narrow", sans-serif;
            font-size: 25px;
            letter-spacing: -0.03em;
            line-height: 1.15;
        }
        .section {
            margin-top: 26px;
        }
        .section h2 {
            margin: 0 0 8px;
            font-family: "Franklin Gothic Medium", "Arial Narrow", sans-serif;
            font-size: 30px;
            letter-spacing: -0.03em;
        }
        .section p.head {
            margin: 0 0 16px;
            color: var(--muted);
            max-width: 860px;
            line-height: 1.56;
        }
        .responsible-grid {
            display: grid;
            gap: 18px;
        }
        .responsible-card {
            display: grid;
            grid-template-columns: 1.05fr 0.95fr;
            gap: 18px;
            padding: 24px;
            border-radius: 30px;
            border: 1px solid var(--line);
            background: var(--panel-strong);
            box-shadow: var(--shadow);
        }
        .responsible-box,
        .entry-box {
            padding: 20px;
            border-radius: 22px;
            border: 1px solid rgba(22, 35, 43, 0.10);
        }
        .responsible-box {
            background: rgba(246, 249, 247, 0.85);
        }
        .entry-box {
            background: linear-gradient(135deg, rgba(15, 111, 103, 0.08), rgba(190, 116, 51, 0.08));
        }
        .responsible-box h3,
        .entry-box h3 {
            margin: 12px 0 8px;
            font-family: "Franklin Gothic Medium", "Arial Narrow", sans-serif;
            font-size: 32px;
            letter-spacing: -0.04em;
            line-height: 0.98;
        }
        .responsible-owner {
            color: var(--accent);
            font-weight: 700;
            margin-bottom: 10px;
        }
        .responsible-box p,
        .entry-box p {
            margin: 0;
            color: var(--muted);
            line-height: 1.58;
        }
        .tag {
            display: inline-flex;
            align-items: center;
            padding: 6px 10px;
            border-radius: 999px;
            background: rgba(15, 111, 103, 0.10);
            color: var(--accent);
            font-size: 12px;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.07em;
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
        .window-preview {
            margin-top: 14px;
            display: grid;
            gap: 12px;
        }
        .window-chip {
            display: flex;
            justify-content: space-between;
            align-items: center;
            gap: 12px;
            padding: 14px 16px;
            border-radius: 18px;
            border: 1px solid rgba(22, 35, 43, 0.12);
            background: rgba(255, 255, 255, 0.74);
        }
        .window-chip strong {
            display: block;
            font-size: 18px;
            line-height: 1.2;
        }
        .window-chip span {
            color: var(--muted);
            font-size: 14px;
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
            .stats,
            .responsible-card {
                grid-template-columns: 1fr;
            }
            .window-chip {
                align-items: flex-start;
                flex-direction: column;
            }
        }
    </style>
</head>
<body>
<div class="page">
    <div class="topbar">
        <div>
            <span class="chip"><?= htmlspecialchars($brand['platform']) ?></span>
            <strong><?= htmlspecialchars($brand['company']) ?></strong>
            <span class="note"><?= htmlspecialchars($brand['tagline']) ?></span>
        </div>
        <span class="status-pill">Entorno de produccion interno</span>
    </div>

    <section class="hero">
        <div class="hero-grid">
            <div>
                <span class="chip">Portal principal</span>
                <h1>Ingreso por responsable y por ventana operativa.</h1>
                <p class="lead">
                    El portal queda organizado para que cada persona entre a su proceso, luego a su ventana de trabajo
                    y desde ahi ejecute las funciones asignadas sin mezclar acciones de otras areas.
                </p>
                <div class="hero-note">Acceso LAN: <?= htmlspecialchars($publicRootUrl) ?></div>
            </div>
            <?php if ($brand['logo_exists']): ?>
                <div class="logo-stage">
                    <div class="logo-frame">
                        <img src="<?= htmlspecialchars($brand['logo_url']) ?>" alt="<?= htmlspecialchars($brand['company']) ?>">
                    </div>
                </div>
            <?php endif; ?>
        </div>
    </section>

    <section class="stats">
        <article class="stat">
            <span>Responsables activos</span>
            <strong><?= count($workspaces) ?></strong>
        </article>
        <article class="stat">
            <span>Ventanas operativas</span>
            <strong><?= $windowCount ?></strong>
        </article>
        <article class="stat">
            <span>Procesos disponibles</span>
            <strong><?= $activeCount + $bundleCount ?></strong>
        </article>
        <article class="stat">
            <span>Salida publica</span>
            <strong>.xls y .xlsx</strong>
        </article>
    </section>

    <section class="section">
        <h2>Procesos Responsables</h2>
        <p class="head">
            Cada responsable entra a su area y desde ahi navega a sus ventanas operativas. Las acciones internas se mantienen intactas.
        </p>

        <div class="responsible-grid">
            <?php foreach ($workspaces as $workspace): ?>
                <article class="responsible-card">
                    <div class="responsible-box">
                        <span class="tag">Proceso responsable</span>
                        <h3><?= htmlspecialchars($workspace['title']) ?></h3>
                        <div class="responsible-owner"><?= htmlspecialchars($workspace['person']) ?></div>
                        <p><?= htmlspecialchars($workspace['summary']) ?></p>
                        <div class="trace">
                            <strong>Ruta de salida:</strong> `storage/outputs`.<br>
                            <strong>Descarga publica:</strong> `.xls` y `.xlsx` segun proceso.<br>
                            <strong>Auditoria:</strong> archivo interno para control del area.<br>
                            <strong>Acceso LAN:</strong> <a href="<?= htmlspecialchars((string)($workspace['home_url'] ?? $publicRootUrl)) ?>">Ingresar</a>
                        </div>
                    </div>

                    <div class="entry-box">
                        <span class="tag">Ventanas activas</span>
                        <h3><?= count($workspace['windows'] ?? []) ?> ventana operativa</h3>
                        <p><?= htmlspecialchars($workspace['future_note']) ?></p>
                        <div class="window-preview">
                            <?php foreach (($workspace['windows'] ?? []) as $window): ?>
                                <div class="window-chip">
                                    <div>
                                        <strong><?= htmlspecialchars($window['title']) ?></strong>
                                        <span><?= htmlspecialchars($window['summary']) ?></span>
                                    </div>
                                </div>
                            <?php endforeach; ?>
                        </div>
                        <a class="cta" href="<?= htmlspecialchars((string)($workspace['home_url'] ?? app_url())) ?>">Ingresar</a>
                    </div>
                </article>
            <?php endforeach; ?>
        </div>
    </section>
</div>
</body>
</html>
