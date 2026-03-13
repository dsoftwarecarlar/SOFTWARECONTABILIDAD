<?php
declare(strict_types=1);

/** @var array<string, mixed> $brand */
/** @var array<string, mixed> $workspace */
/** @var array<string, mixed> $currentModule */
/** @var array<string, mixed>|null $result */
/** @var string|null $error */
/** @var array<int, array<string, mixed>> $history */
/** @var array<string, mixed> $pageConfig */
?>
<!doctype html>
<html lang="es">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title><?= htmlspecialchars((string)$workspace['title']) ?> | <?= htmlspecialchars((string)$currentModule['title']) ?></title>
    <style>
        :root {
            --ink: #18242c;
            --muted: #5f6970;
            --line: rgba(24, 36, 44, 0.11);
            --panel: rgba(255, 252, 246, 0.92);
            --deep: #112530;
            --accent: #106f66;
            --bg-1: #ece3d3;
            --bg-2: #e5eded;
            --ok-bg: #e6f4ea;
            --ok-ink: #25573e;
            --err-bg: #fae7e3;
            --err-ink: #7e2f27;
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
        .page { max-width: 1280px; margin: 0 auto; padding: 28px 18px 44px; }
        .topbar { display: flex; justify-content: space-between; align-items: center; gap: 14px; margin-bottom: 16px; }
        .chip { display: inline-flex; align-items: center; padding: 6px 10px; border-radius: 999px; background: rgba(16, 111, 102, 0.10); color: var(--accent); font-size: 12px; font-weight: 700; letter-spacing: 0.08em; text-transform: uppercase; }
        .topbar strong { display: block; margin-top: 8px; font-family: "Franklin Gothic Medium", "Arial Narrow", sans-serif; font-size: 28px; letter-spacing: -0.03em; }
        .topbar span.note { display: block; margin-top: 3px; color: var(--muted); font-size: 14px; }
        .back-link { color: var(--deep); text-decoration: none; font-weight: 700; }
        .hero { position: relative; overflow: hidden; border-radius: 34px; border: 1px solid var(--line); box-shadow: var(--shadow); background: linear-gradient(135deg, rgba(255, 252, 246, 0.98), rgba(243, 235, 220, 0.84)); padding: 30px; }
        .hero::after { content: ""; position: absolute; right: -68px; top: -80px; width: 290px; height: 290px; border-radius: 50%; background: radial-gradient(circle, rgba(16, 111, 102, 0.16), transparent 68%); }
        .hero-grid { position: relative; z-index: 1; display: grid; grid-template-columns: minmax(0, 1fr) 320px; gap: 22px; align-items: center; }
        h1 { margin: 14px 0 12px; font-family: "Bodoni MT", "Book Antiqua", serif; font-size: clamp(38px, 5.8vw, 72px); line-height: 0.92; letter-spacing: -0.05em; max-width: 780px; }
        .lead { margin: 0; color: var(--muted); max-width: 760px; font-size: 17px; line-height: 1.62; }
        .hero-note { margin-top: 18px; display: inline-flex; align-items: center; padding: 10px 14px; border-radius: 16px; background: rgba(16, 111, 102, 0.09); color: var(--deep); font-weight: 700; }
        .hero-side { display: grid; gap: 12px; }
        .logo-stage { padding: 16px; border-radius: 28px; background: linear-gradient(180deg, rgba(21, 38, 48, 0.95), rgba(14, 28, 37, 0.96)); box-shadow: inset 0 0 0 1px rgba(255, 255, 255, 0.07); display: grid; place-items: center; }
        .logo-frame { width: 180px; height: 180px; border-radius: 38px; display: grid; place-items: center; background: linear-gradient(180deg, rgba(255, 255, 255, 0.98), rgba(248, 239, 224, 0.95)); box-shadow: 0 18px 34px rgba(0, 0, 0, 0.22); }
        .logo-frame img { width: 136px; height: 136px; object-fit: contain; }
        .hero-card { padding: 16px 18px; border-radius: 22px; background: rgba(255, 255, 255, 0.76); border: 1px solid rgba(24, 36, 44, 0.10); }
        .hero-card span, .map-item span, .module-link span, .stats-item span { display: block; font-size: 12px; letter-spacing: 0.08em; text-transform: uppercase; color: var(--muted); font-weight: 700; }
        .hero-card strong, .map-item strong, .module-link strong, .stats-item strong { display: block; margin-top: 8px; font-size: 18px; line-height: 1.25; }
        .layout { margin-top: 18px; display: grid; grid-template-columns: 360px minmax(0, 1fr); gap: 16px; align-items: start; }
        .panel { border-radius: 28px; border: 1px solid var(--line); background: var(--panel); box-shadow: var(--shadow); padding: 22px; }
        .panel h2 { margin: 0 0 8px; font-family: "Franklin Gothic Medium", "Arial Narrow", sans-serif; font-size: 25px; letter-spacing: -0.03em; }
        .panel p { margin: 0; color: var(--muted); line-height: 1.58; }
        .side-stack, .main-stack, .history, .module-links, .stats { display: grid; gap: 12px; }
        .map-item, .module-link, .stats-item, .history-item { padding: 16px 18px; border-radius: 20px; border: 1px solid rgba(24, 36, 44, 0.10); background: rgba(255, 255, 255, 0.78); }
        .map-item p, .module-link p, .stats-item p { margin-top: 8px; }
        .map-item a { display: inline-block; margin-top: 10px; color: var(--deep); font-weight: 700; text-decoration: none; }
        .module-link { text-decoration: none; color: inherit; }
        .form-grid { margin-top: 14px; display: grid; grid-template-columns: minmax(0, 1fr) 300px; gap: 16px; align-items: start; }
        form { display: grid; gap: 12px; }
        label.file-label { display: grid; gap: 6px; font-size: 13px; color: var(--deep); font-weight: 700; }
        input[type="file"] { width: 100%; padding: 12px; border-radius: 14px; border: 1px dashed rgba(16, 111, 102, 0.34); background: rgba(255, 255, 255, 0.75); }
        button, .button-link { display: inline-flex; align-items: center; justify-content: center; padding: 12px 18px; border-radius: 999px; border: 0; background: linear-gradient(135deg, var(--accent) 0%, var(--deep) 100%); color: #fff; text-decoration: none; font-weight: 700; cursor: pointer; }
        .msg { margin-top: 14px; padding: 14px 16px; border-radius: 18px; white-space: pre-wrap; }
        .msg.ok { background: var(--ok-bg); color: var(--ok-ink); }
        .msg.err { background: var(--err-bg); color: var(--err-ink); }
        .console { margin-top: 12px; padding: 14px; border-radius: 18px; background: #18242c; color: #dfe7e4; font-family: Consolas, "Courier New", monospace; font-size: 12px; max-height: 240px; overflow: auto; }
        .history-item { display: grid; grid-template-columns: 1fr auto; align-items: center; gap: 12px; }
        .history-item strong { display: block; }
        .history-meta { margin-top: 6px; color: var(--muted); font-size: 13px; }
        @media (max-width: 980px) {
            .hero-grid, .layout, .form-grid, .history-item { grid-template-columns: 1fr; }
        }
    </style>
</head>
<body>
<div class="page">
    <div class="topbar">
        <div>
            <span class="chip"><?= htmlspecialchars((string)$brand['platform']) ?></span>
            <strong><?= htmlspecialchars((string)($pageConfig['top_title'] ?? $currentModule['title'])) ?></strong>
            <span class="note"><?= htmlspecialchars((string)($pageConfig['top_note'] ?? '')) ?></span>
        </div>
        <a class="back-link" href="<?= htmlspecialchars((string)($pageConfig['back_url'] ?? app_url())) ?>"><?= htmlspecialchars((string)($pageConfig['back_label'] ?? 'Volver')) ?></a>
    </div>

    <section class="hero">
        <div class="hero-grid">
            <div>
                <span class="chip"><?= htmlspecialchars((string)($pageConfig['hero_chip'] ?? 'Proceso')) ?></span>
                <h1><?= htmlspecialchars((string)($pageConfig['hero_title'] ?? $currentModule['title'])) ?></h1>
                <p class="lead"><?= htmlspecialchars((string)($pageConfig['hero_lead'] ?? '')) ?></p>
                <div class="hero-note"><?= htmlspecialchars((string)($pageConfig['hero_note'] ?? '')) ?></div>
            </div>
            <div class="hero-side">
                <?php if (!empty($pageConfig['show_logo']) && !empty($brand['logo_exists'])): ?>
                    <div class="logo-stage">
                        <div class="logo-frame">
                            <img src="<?= htmlspecialchars((string)$brand['logo_url']) ?>" alt="<?= htmlspecialchars((string)$brand['company']) ?>">
                        </div>
                    </div>
                <?php endif; ?>
                <div class="hero-card">
                    <span><?= htmlspecialchars((string)($pageConfig['hero_card_label'] ?? 'Resumen')) ?></span>
                    <strong><?= htmlspecialchars((string)($pageConfig['hero_card_value'] ?? '')) ?></strong>
                </div>
            </div>
        </div>
    </section>

    <div class="layout">
        <aside class="side-stack">
            <?php foreach ((array)($pageConfig['sidebar_panels'] ?? []) as $panel): ?>
                <section class="panel">
                    <h2><?= htmlspecialchars((string)($panel['title'] ?? 'Panel')) ?></h2>
                    <?php if (!empty($panel['description'])): ?>
                        <p><?= htmlspecialchars((string)$panel['description']) ?></p>
                    <?php endif; ?>
                    <?php if (($panel['type'] ?? 'map') === 'links'): ?>
                        <div class="module-links" style="margin-top: 14px;">
                            <?php foreach ((array)($panel['items'] ?? []) as $item): ?>
                                <a class="module-link" href="<?= htmlspecialchars((string)($item['url'] ?? '#')) ?>">
                                    <strong><?= htmlspecialchars((string)($item['title'] ?? '')) ?></strong>
                                    <p><?= htmlspecialchars((string)($item['description'] ?? '')) ?></p>
                                </a>
                            <?php endforeach; ?>
                        </div>
                    <?php else: ?>
                        <?php foreach ((array)($panel['items'] ?? []) as $item): ?>
                            <div class="map-item" style="margin-top: 12px;">
                                <span><?= htmlspecialchars((string)($item['tag'] ?? 'Dato')) ?></span>
                                <strong><?= htmlspecialchars((string)($item['title'] ?? '')) ?></strong>
                                <p><?= htmlspecialchars((string)($item['description'] ?? '')) ?></p>
                                <?php if (!empty($item['url']) && !empty($item['link_label'])): ?>
                                    <a href="<?= htmlspecialchars((string)$item['url']) ?>"><?= htmlspecialchars((string)$item['link_label']) ?></a>
                                <?php endif; ?>
                            </div>
                        <?php endforeach; ?>
                    <?php endif; ?>
                </section>
            <?php endforeach; ?>
        </aside>

        <main class="main-stack">
            <section class="panel">
                <h2><?= htmlspecialchars((string)($pageConfig['upload_title'] ?? 'Procesar archivo')) ?></h2>
                <p><?= htmlspecialchars((string)($pageConfig['upload_description'] ?? '')) ?></p>

                <div class="form-grid">
                    <div>
                        <form method="post" enctype="multipart/form-data">
                            <label class="file-label">
                                <?= htmlspecialchars((string)($pageConfig['form_label'] ?? 'Archivo')) ?>
                                <input
                                    type="file"
                                    name="<?= htmlspecialchars((string)($pageConfig['form_field_name'] ?? 'file')) ?>"
                                    accept="<?= htmlspecialchars((string)($pageConfig['form_accept'] ?? '')) ?>"
                                    <?= !empty($pageConfig['form_multiple']) ? 'multiple' : '' ?>
                                    required
                                >
                            </label>
                            <button type="submit"><?= htmlspecialchars((string)($pageConfig['button_label'] ?? 'Procesar')) ?></button>
                        </form>

                        <?php if ($error !== null): ?>
                            <div class="msg err"><?= htmlspecialchars($error) ?></div>
                        <?php endif; ?>

                        <?php if ($result !== null): ?>
                            <div class="msg ok"><?= htmlspecialchars((string)($pageConfig['success_message'] ?? ('Archivo generado correctamente: ' . ($result['excel_name'] ?? '')))) ?></div>
                            <a class="button-link" href="<?= htmlspecialchars((string)$result['download_url']) ?>"><?= htmlspecialchars((string)($pageConfig['download_label'] ?? 'Descargar')) ?></a>
                            <?php if (($result['console'] ?? '') !== ''): ?>
                                <pre class="console"><?= htmlspecialchars((string)$result['console']) ?></pre>
                            <?php endif; ?>
                        <?php endif; ?>
                    </div>

                    <div class="stats">
                        <?php foreach ((array)($pageConfig['stats_items'] ?? []) as $item): ?>
                            <div class="stats-item">
                                <span><?= htmlspecialchars((string)($item['tag'] ?? 'Dato')) ?></span>
                                <strong><?= htmlspecialchars((string)($item['title'] ?? '')) ?></strong>
                                <p><?= htmlspecialchars((string)($item['description'] ?? '')) ?></p>
                                <?php if (!empty($item['url']) && !empty($item['button_label'])): ?>
                                    <a class="button-link" style="margin-top: 10px;" href="<?= htmlspecialchars((string)$item['url']) ?>"><?= htmlspecialchars((string)$item['button_label']) ?></a>
                                <?php endif; ?>
                            </div>
                        <?php endforeach; ?>
                    </div>
                </div>
            </section>

            <?php foreach ((array)($pageConfig['secondary_panels'] ?? []) as $panel): ?>
                <section class="panel">
                    <h2><?= htmlspecialchars((string)($panel['title'] ?? 'Detalle')) ?></h2>
                    <p><?= htmlspecialchars((string)($panel['description'] ?? '')) ?></p>
                    <div class="history" style="margin-top: 14px;">
                        <?php foreach ((array)($panel['items'] ?? []) as $item): ?>
                            <div class="history-item">
                                <div>
                                    <strong><?= htmlspecialchars((string)($item['title'] ?? '')) ?></strong>
                                    <div class="history-meta"><?= htmlspecialchars((string)($item['meta'] ?? '')) ?></div>
                                </div>
                                <?php if (!empty($item['url']) && !empty($item['button_label'])): ?>
                                    <div>
                                        <a class="button-link" href="<?= htmlspecialchars((string)$item['url']) ?>"><?= htmlspecialchars((string)$item['button_label']) ?></a>
                                    </div>
                                <?php endif; ?>
                            </div>
                        <?php endforeach; ?>
                    </div>
                </section>
            <?php endforeach; ?>

            <section class="panel">
                <h2><?= htmlspecialchars((string)($pageConfig['history_title'] ?? 'Historial reciente')) ?></h2>
                <p><?= htmlspecialchars((string)($pageConfig['history_description'] ?? '')) ?></p>
                <div class="history" style="margin-top: 14px;">
                    <?php if ($history === []): ?>
                        <div class="history-item">
                            <div>
                                <strong><?= htmlspecialchars((string)($pageConfig['history_empty_title'] ?? 'Sin archivos')) ?></strong>
                                <div class="history-meta"><?= htmlspecialchars((string)($pageConfig['history_empty_description'] ?? '')) ?></div>
                            </div>
                        </div>
                    <?php else: ?>
                        <?php foreach ($history as $item): ?>
                            <div class="history-item">
                                <div>
                                    <strong><?= htmlspecialchars((string)$item['name']) ?></strong>
                                    <div class="history-meta"><?= htmlspecialchars((string)$item['time']) ?> | <?= number_format((int)$item['size'] / 1024, 2) ?> KB</div>
                                </div>
                                <div>
                                    <a class="button-link" href="<?= htmlspecialchars(app_output_download_url((string)$item['name'])) ?>">Descargar</a>
                                </div>
                            </div>
                        <?php endforeach; ?>
                    <?php endif; ?>
                </div>
            </section>
        </main>
    </div>
</div>
</body>
</html>
