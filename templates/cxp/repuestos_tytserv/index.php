<?php
declare(strict_types=1);

/** @var array<string, mixed> $brand */
/** @var array<string, mixed> $workspace */
/** @var array<string, mixed> $currentModule */
/** @var array<string, mixed>|null $window */
/** @var string $templateDir */
/** @var string $templateFileName */
/** @var array<int, array<string, string>> $fileFields */
/** @var array<string, mixed>|null $result */
/** @var string|null $error */
/** @var array<int, array<string, mixed>> $history */
/** @var array<string, string> $pageConfig */
?>
<!doctype html>
<html lang="es">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title><?= htmlspecialchars((string)$workspace['title']) ?> | <?= htmlspecialchars((string)$currentModule['title']) ?></title>
    <style>
        :root {
            --ink: #17242d;
            --muted: #5f6b72;
            --line: rgba(23, 36, 45, 0.12);
            --panel: rgba(255, 252, 246, 0.92);
            --deep: #102430;
            --accent: #0f6f67;
            --bg-1: #ece4d5;
            --bg-2: #e4edef;
            --ok-bg: #e6f5ec;
            --ok-ink: #20553d;
            --err-bg: #fae7e3;
            --err-ink: #7f3028;
            --shadow: 0 20px 48px rgba(16, 27, 36, 0.10);
        }
        * { box-sizing: border-box; }
        body {
            margin: 0;
            color: var(--ink);
            background:
                radial-gradient(circle at 12% 12%, rgba(191, 117, 52, 0.22), transparent 23%),
                radial-gradient(circle at 86% 11%, rgba(15, 111, 103, 0.18), transparent 24%),
                linear-gradient(180deg, var(--bg-1) 0%, var(--bg-2) 100%);
            font-family: "Trebuchet MS", "Lucida Sans Unicode", sans-serif;
        }
        .page { max-width: 1280px; margin: 0 auto; padding: 28px 18px 44px; }
        .topbar { display: flex; justify-content: space-between; align-items: flex-start; flex-wrap: wrap; gap: 14px; margin-bottom: 16px; }
        .topbar > div { min-width: 0; flex: 1 1 420px; }
        .chip { display: inline-flex; align-items: center; padding: 6px 10px; border-radius: 999px; background: rgba(15,111,103,.1); color: var(--accent); font-size: 12px; font-weight: 700; letter-spacing: .08em; text-transform: uppercase; }
        .topbar strong { display: block; margin-top: 8px; font-family: "Franklin Gothic Medium", "Arial Narrow", sans-serif; font-size: 28px; letter-spacing: -.03em; }
        .topbar span.note { display: block; margin-top: 3px; color: var(--muted); font-size: 14px; }
        .back-link { color: var(--deep); text-decoration: none; font-weight: 700; margin-left: auto; }
        .hero, .panel { border-radius: 28px; border: 1px solid var(--line); box-shadow: var(--shadow); }
        .hero {
            background: linear-gradient(135deg, rgba(255,252,246,.98), rgba(243,235,220,.84));
            padding: 30px;
            position: relative;
            overflow: hidden;
        }
        .hero::after {
            content: "";
            position: absolute;
            right: -70px;
            top: -76px;
            width: 280px;
            height: 280px;
            border-radius: 50%;
            background: radial-gradient(circle, rgba(15,111,103,.16), transparent 68%);
        }
        .hero-grid { position: relative; z-index: 1; display: grid; grid-template-columns: minmax(0,1fr) minmax(250px, 320px); gap: 22px; align-items: start; }
        h1 { margin: 14px 0 12px; font-family: "Bodoni MT", "Book Antiqua", serif; font-size: clamp(38px,5.8vw,72px); line-height: .92; letter-spacing: -.05em; max-width: 780px; }
        .lead, .panel p { margin: 0; color: var(--muted); line-height: 1.6; }
        .hero-note { margin-top: 18px; display: inline-flex; padding: 10px 14px; border-radius: 16px; background: rgba(16,111,102,.09); color: var(--deep); font-weight: 700; }
        .hero-card { padding: 18px; border-radius: 22px; background: rgba(255,255,255,.76); border: 1px solid rgba(23,36,45,.1); }
        .hero-card span { display: block; font-size: 12px; color: var(--muted); font-weight: 700; letter-spacing: .08em; text-transform: uppercase; }
        .hero-card strong { display: block; margin-top: 8px; font-size: 18px; line-height: 1.25; overflow-wrap: anywhere; }
        .layout { margin-top: 18px; display: grid; grid-template-columns: minmax(300px, 360px) minmax(0,1fr); gap: 16px; align-items: start; }
        .panel { background: var(--panel); padding: 22px; }
        .panel h2 { margin: 0 0 8px; font-family: "Franklin Gothic Medium", "Arial Narrow", sans-serif; font-size: 25px; letter-spacing: -.03em; }
        .stack, .history, .summary-grid { display: grid; gap: 12px; }
        .box { padding: 16px 18px; border-radius: 20px; border: 1px solid rgba(24,36,44,.1); background: rgba(255,255,255,.78); }
        .box strong { display: block; margin-top: 8px; font-size: 18px; line-height: 1.25; overflow-wrap: anywhere; }
        .tag { display: block; font-size: 12px; color: var(--muted); font-weight: 700; letter-spacing: .08em; text-transform: uppercase; }
        .meta { margin-top: 6px; color: var(--muted); font-size: 13px; }
        .base-path { margin-top: 10px; font-size: 12px; color: var(--muted); word-break: break-all; }
        .form-grid { display: grid; grid-template-columns: minmax(0,1fr) 300px; gap: 16px; margin-top: 14px; align-items: start; }
        form { display: grid; gap: 12px; }
        input[type="file"] { width: 100%; padding: 12px; border-radius: 14px; border: 1px dashed rgba(16,111,102,.34); background: rgba(255,255,255,.75); }
        label.file-label { display: grid; gap: 6px; font-size: 13px; color: var(--deep); font-weight: 700; }
        button, .button-link { display: inline-flex; align-items: center; justify-content: center; padding: 12px 18px; border-radius: 999px; border: 0; background: linear-gradient(135deg, var(--accent) 0%, var(--deep) 100%); color: #fff; text-decoration: none; font-weight: 700; cursor: pointer; }
        .msg { margin-top: 14px; padding: 14px 16px; border-radius: 18px; white-space: pre-wrap; }
        .msg.ok { background: var(--ok-bg); color: var(--ok-ink); }
        .msg.err { background: var(--err-bg); color: var(--err-ink); }
        .console { margin-top: 12px; padding: 14px; border-radius: 18px; background: #18242c; color: #dfe7e4; font-family: Consolas, "Courier New", monospace; font-size: 12px; max-height: 240px; overflow: auto; }
        .history-item { display: flex; justify-content: space-between; align-items: center; gap: 12px; }
        .history-item > div { min-width: 0; }
        @media (max-width: 1200px) {
            .layout { grid-template-columns: 1fr; }
        }
        @media (max-width: 980px) {
            .hero-grid, .form-grid { grid-template-columns: 1fr; }
            .history-item { flex-direction: column; align-items: flex-start; }
            .back-link { margin-left: 0; }
            .hero { padding: 22px; }
        }
    </style>
</head>
<body>
<div class="page">
    <div class="topbar">
        <div>
            <span class="chip"><?= htmlspecialchars((string)$brand['platform']) ?></span>
            <strong><?= htmlspecialchars((string)$currentModule['title']) ?></strong>
            <span class="note"><?= htmlspecialchars((string)($pageConfig['top_note'] ?? '')) ?></span>
        </div>
        <a class="back-link" href="<?= htmlspecialchars((string)($window['url'] ?? ($workspace['home_url'] ?? app_url()))) ?>">Volver a la ventana</a>
    </div>

    <section class="hero">
        <div class="hero-grid">
            <div>
                <span class="chip"><?= htmlspecialchars((string)($pageConfig['hero_chip'] ?? 'Proceso')) ?></span>
                <h1><?= htmlspecialchars((string)$currentModule['title']) ?></h1>
                <p class="lead"><?= htmlspecialchars((string)($pageConfig['hero_lead'] ?? '')) ?></p>
                <div class="hero-note"><?= htmlspecialchars((string)($pageConfig['hero_note'] ?? '')) ?></div>
            </div>
            <div class="hero-card">
                <span>Plantilla base</span>
                <strong><?= htmlspecialchars($templateFileName) ?></strong>
                <div class="base-path"><?= htmlspecialchars($templateDir) ?></div>
            </div>
        </div>
    </section>

    <div class="layout">
        <aside class="panel stack">
            <h2><?= htmlspecialchars((string)($pageConfig['inputs_title'] ?? 'Entradas requeridas')) ?></h2>
            <?php foreach ($fileFields as $fieldConfig): ?>
                <article class="box">
                    <span class="tag">Archivo</span>
                    <strong><?= htmlspecialchars((string)($fieldConfig['label'] ?? 'Archivo')) ?></strong>
                </article>
            <?php endforeach; ?>
            <article class="box">
                <span class="tag"><?= htmlspecialchars((string)($pageConfig['retention_tag'] ?? 'Retencion')) ?></span>
                <strong><?= app_output_retention_limit() ?> ultimos reportes</strong>
                <div class="meta"><?= htmlspecialchars((string)($pageConfig['retention_description'] ?? '')) ?></div>
            </article>
        </aside>

        <main class="stack">
            <section class="panel">
                <h2><?= htmlspecialchars((string)($pageConfig['process_title'] ?? 'Procesar plantilla')) ?></h2>
                <p><?= htmlspecialchars((string)($pageConfig['process_description'] ?? '')) ?></p>

                <div class="form-grid">
                    <div>
                        <form method="post" enctype="multipart/form-data">
                            <?php foreach ($fileFields as $fieldConfig): ?>
                                <label class="file-label">
                                    <?= htmlspecialchars((string)($fieldConfig['label'] ?? 'Archivo')) ?>
                                    <input type="file" name="<?= htmlspecialchars((string)($fieldConfig['field'] ?? 'archivo')) ?>" accept=".xls,.xlsx" required>
                                </label>
                            <?php endforeach; ?>
                            <button type="submit"><?= htmlspecialchars((string)($pageConfig['button_label'] ?? 'Procesar')) ?></button>
                        </form>

                        <?php if ($error !== null): ?>
                            <div class="msg err"><?= htmlspecialchars($error) ?></div>
                        <?php endif; ?>

                        <?php if ($result !== null): ?>
                            <div class="msg ok">Archivo generado correctamente: <?= htmlspecialchars((string)$result['excel_name']) ?></div>
                            <a class="button-link" href="<?= htmlspecialchars((string)$result['download_url']) ?>">Descargar reporte final</a>
                            <?php if (($result['summary'] ?? []) !== []): ?>
                                <div class="summary-grid" style="margin-top: 12px;">
                                    <?php foreach ((array)$result['summary'] as $item): ?>
                                        <article class="box">
                                            <span class="tag"><?= htmlspecialchars((string)($item['label'] ?? 'RESUMEN')) ?></span>
                                            <strong><?= (int)($item['rows'] ?? 0) ?> filas procesadas</strong>
                                        </article>
                                    <?php endforeach; ?>
                                </div>
                            <?php endif; ?>
                            <?php if (($result['console'] ?? '') !== ''): ?>
                                <pre class="console"><?= htmlspecialchars((string)$result['console']) ?></pre>
                            <?php endif; ?>
                        <?php endif; ?>
                    </div>

                    <article class="box">
                        <span class="tag"><?= htmlspecialchars((string)($pageConfig['output_tag'] ?? 'Salida')) ?></span>
                        <strong><?= htmlspecialchars((string)($pageConfig['output_title'] ?? 'storage/outputs')) ?></strong>
                        <div class="meta"><?= htmlspecialchars((string)($pageConfig['output_description'] ?? '')) ?></div>
                    </article>
                </div>
            </section>

            <section class="panel">
                <h2><?= htmlspecialchars((string)($pageConfig['history_title'] ?? 'Historial')) ?></h2>
                <p>Se conservan solo los ultimos <?= app_output_retention_limit() ?> reportes de esta accion.</p>
                <div class="history" style="margin-top: 14px;">
                    <?php if ($history === []): ?>
                        <article class="box">
                            <strong><?= htmlspecialchars((string)($pageConfig['history_empty_title'] ?? 'Sin historial')) ?></strong>
                            <div class="meta"><?= htmlspecialchars((string)($pageConfig['history_empty_description'] ?? '')) ?></div>
                        </article>
                    <?php else: ?>
                        <?php foreach ($history as $item): ?>
                            <article class="box history-item">
                                <div>
                                    <strong><?= htmlspecialchars((string)$item['name']) ?></strong>
                                    <div class="meta"><?= htmlspecialchars((string)$item['time']) ?> - <?= number_format(((int)$item['size']) / 1024, 1) ?> KB</div>
                                </div>
                                <a class="button-link" href="<?= htmlspecialchars(app_output_download_url((string)$item['name'])) ?>">Descargar</a>
                            </article>
                        <?php endforeach; ?>
                    <?php endif; ?>
                </div>
            </section>
        </main>
    </div>
</div>
</body>
</html>
