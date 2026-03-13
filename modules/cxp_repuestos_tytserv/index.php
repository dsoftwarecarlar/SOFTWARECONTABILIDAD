<?php
declare(strict_types=1);

require dirname(__DIR__, 2) . '/includes/app.php';

$brand = app_brand();
$context = app_workspace_module('cxp', 'cxp_repuestos_tytserv');
if ($context === null) {
    throw new RuntimeException('No se encontro la configuracion del modulo.');
}

$windowContext = app_workspace_window('cxp', 'facturacion_repuestos_tytserv');
$workspace = $context['workspace'];
$currentModule = $context['module'];
$window = $windowContext['window'] ?? null;
$uploadsDir = app_ensure_dir(app_storage_path('uploads'));
$outputsDir = app_ensure_dir(app_storage_path('outputs'));
$templateDir = app_join_path(app_root(), 'outputs', 'EJEMPLOAMANOTAREA3');
$templatePath = app_join_path($templateDir, 'FACTURACION REPUESTOS TYTSERV FEBRERO 2026.xlsx');
$requestMethod = $_SERVER['REQUEST_METHOD'] ?? 'GET';

$result = null;
$error = null;

$fileFields = [
    'excel_tyt' => 'MATRIZ (RepLibroVentasGeneral)',
    'excel_peug' => 'PEUGEOT (RepLibroVentasGeneral)',
    'excel_chgn' => 'CHANGAN (RepLibroVentasGeneral)',
    'excel_szk' => 'SUZUKI (RepLibroVentasGeneral)',
];

if ($requestMethod === 'POST') {
    try {
        if (!is_file($templatePath)) {
            throw new RuntimeException('No existe la plantilla base FACTURACION REPUESTOS TYTSERV FEBRERO 2026.xlsx.');
        }

        $savedInputs = [];
        $stamp = date('Ymd_His');

        foreach ($fileFields as $field => $label) {
            if (!isset($_FILES[$field])) {
                throw new RuntimeException("No se recibio el archivo requerido: $label.");
            }

            $file = $_FILES[$field];
            if (!is_array($file) || (int)($file['error'] ?? UPLOAD_ERR_NO_FILE) !== UPLOAD_ERR_OK) {
                throw new RuntimeException("Error al subir el archivo: $label.");
            }

            $originalName = trim((string)($file['name'] ?? 'archivo.xlsx'));
            $tmpPath = (string)($file['tmp_name'] ?? '');
            $extension = strtolower(pathinfo($originalName, PATHINFO_EXTENSION));
            if (!in_array($extension, ['xls', 'xlsx'], true)) {
                throw new RuntimeException("Formato no permitido en $label. Solo .xls o .xlsx.");
            }

            $safeBase = preg_replace('/[^A-Za-z0-9_-]+/', '_', pathinfo($originalName, PATHINFO_FILENAME));
            $safeBase = trim((string)$safeBase, '_');
            if ($safeBase === '') {
                $safeBase = $field;
            }

            $inputName = sprintf('%s_%s_%s.%s', $field, $safeBase, $stamp, $extension);
            $inputPath = app_join_path($uploadsDir, $inputName);
            if (!move_uploaded_file($tmpPath, $inputPath)) {
                throw new RuntimeException("No se pudo guardar el archivo subido: $label.");
            }

            $savedInputs[$field] = $inputPath;
        }

        $outputName = 'repuestos_tytserv_' . $stamp . '.xlsx';
        $outputPath = app_join_path($outputsDir, $outputName);
        $scriptPath = app_join_path(app_root(), 'run_repuestos_tytserv.ps1');
        if (!is_file($scriptPath)) {
            throw new RuntimeException('No existe run_repuestos_tytserv.ps1 en el proyecto.');
        }

        $powershell = app_join_path(
            getenv('WINDIR') ?: 'C:\\Windows',
            'System32',
            'WindowsPowerShell',
            'v1.0',
            'powershell.exe'
        );
        if (!is_file($powershell)) {
            $powershell = 'powershell.exe';
        }

        $command = implode(' ', [
            escapeshellarg($powershell),
            '-NoProfile',
            '-ExecutionPolicy',
            'Bypass',
            '-File',
            escapeshellarg($scriptPath),
            '-InputTyt',
            escapeshellarg($savedInputs['excel_tyt']),
            '-InputPeug',
            escapeshellarg($savedInputs['excel_peug']),
            '-InputChgn',
            escapeshellarg($savedInputs['excel_chgn']),
            '-InputSzk',
            escapeshellarg($savedInputs['excel_szk']),
            '-TemplatePath',
            escapeshellarg($templatePath),
            '-OutputPath',
            escapeshellarg($outputPath),
            '2>&1',
        ]);

        $lines = [];
        $exitCode = 0;
        exec($command, $lines, $exitCode);
        $console = trim(implode(PHP_EOL, array_map(static fn($line): string => trim((string)$line), $lines)));

        if ($exitCode !== 0) {
            throw new RuntimeException($console !== '' ? $console : 'El proceso fallo al generar la plantilla.');
        }

        if (!is_file($outputPath)) {
            throw new RuntimeException('El proceso termino sin generar el archivo de salida.');
        }

        $summary = [];
        foreach ($lines as $line) {
            $line = trim((string)$line);
            if (preg_match('/^INFO\|([a-z0-9_]+)\|rows=(\d+)$/i', $line, $matches) === 1) {
                $summary[strtolower($matches[1])] = (int)$matches[2];
            }
        }

        $orderedSummary = [];
        foreach ([
            'tyt' => 'MATRIZ',
            'peug' => 'PEUGEOT',
            'chgn' => 'CHANGAN',
            'szk' => 'SUZUKI',
        ] as $key => $label) {
            if (!array_key_exists($key, $summary)) {
                continue;
            }

            $orderedSummary[] = [
                'label' => $label,
                'rows' => $summary[$key],
            ];
        }

        $result = [
            'excel_name' => $outputName,
            'download_url' => app_output_download_url($outputName),
            'summary' => $orderedSummary,
            'console' => $console,
        ];

        app_cleanup_output_files_for_action('repuestos_tytserv', app_output_retention_limit());
        app_cleanup_upload_files(app_upload_retention_limit());
    } catch (Throwable $e) {
        $error = $e->getMessage();
    }
}

$history = app_list_output_files_for_action('repuestos_tytserv', app_output_retention_limit());
?>
<!doctype html>
<html lang="es">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title><?= htmlspecialchars($workspace['title']) ?> | <?= htmlspecialchars($currentModule['title']) ?></title>
    <style>
        :root {
            --ink: #17242d;
            --muted: #5f6b72;
            --line: rgba(23, 36, 45, 0.12);
            --panel: rgba(255, 252, 246, 0.92);
            --panel-strong: rgba(255, 255, 255, 0.96);
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
        .topbar { display: flex; justify-content: space-between; align-items: center; gap: 14px; margin-bottom: 16px; }
        .chip { display: inline-flex; align-items: center; padding: 6px 10px; border-radius: 999px; background: rgba(15,111,103,.1); color: var(--accent); font-size: 12px; font-weight: 700; letter-spacing: .08em; text-transform: uppercase; }
        .topbar strong { display: block; margin-top: 8px; font-family: "Franklin Gothic Medium", "Arial Narrow", sans-serif; font-size: 28px; letter-spacing: -.03em; }
        .topbar span.note { display: block; margin-top: 3px; color: var(--muted); font-size: 14px; }
        .back-link { color: var(--deep); text-decoration: none; font-weight: 700; }
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
        .hero-grid { position: relative; z-index: 1; display: grid; grid-template-columns: minmax(0,1fr) 320px; gap: 22px; align-items: center; }
        h1 { margin: 14px 0 12px; font-family: "Bodoni MT", "Book Antiqua", serif; font-size: clamp(38px,5.8vw,72px); line-height: .92; letter-spacing: -.05em; max-width: 780px; }
        .lead, .panel p { margin: 0; color: var(--muted); line-height: 1.6; }
        .hero-note { margin-top: 18px; display: inline-flex; padding: 10px 14px; border-radius: 16px; background: rgba(16,111,102,.09); color: var(--deep); font-weight: 700; }
        .hero-card { padding: 18px; border-radius: 22px; background: rgba(255,255,255,.76); border: 1px solid rgba(23,36,45,.1); }
        .hero-card span { display: block; font-size: 12px; color: var(--muted); font-weight: 700; letter-spacing: .08em; text-transform: uppercase; }
        .hero-card strong { display: block; margin-top: 8px; font-size: 18px; line-height: 1.25; }
        .layout { margin-top: 18px; display: grid; grid-template-columns: 360px minmax(0,1fr); gap: 16px; align-items: start; }
        .panel { background: var(--panel); padding: 22px; }
        .panel h2 { margin: 0 0 8px; font-family: "Franklin Gothic Medium", "Arial Narrow", sans-serif; font-size: 25px; letter-spacing: -.03em; }
        .stack, .history, .summary-grid { display: grid; gap: 12px; }
        .box { padding: 16px 18px; border-radius: 20px; border: 1px solid rgba(24,36,44,.1); background: rgba(255,255,255,.78); }
        .box strong { display: block; margin-top: 8px; font-size: 18px; line-height: 1.25; }
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
        @media (max-width: 980px) {
            .hero-grid, .layout, .form-grid { grid-template-columns: 1fr; }
            .history-item { flex-direction: column; align-items: flex-start; }
        }
    </style>
</head>
<body>
<div class="page">
    <div class="topbar">
        <div>
            <span class="chip"><?= htmlspecialchars($brand['platform']) ?></span>
            <strong><?= htmlspecialchars($currentModule['title']) ?></strong>
            <span class="note">Carga y procesamiento mensual por 4 marcas</span>
        </div>
        <a class="back-link" href="<?= htmlspecialchars((string)($window['url'] ?? ($workspace['home_url'] ?? app_url()))) ?>">Volver a la ventana</a>
    </div>

    <section class="hero">
        <div class="hero-grid">
            <div>
                <span class="chip">Flujo mensual</span>
                <h1><?= htmlspecialchars($currentModule['title']) ?></h1>
                <p class="lead">Carga los 4 archivos RepLibroVentasGeneral del mes y genera la salida final en el formato de la plantilla manual.</p>
                <div class="hero-note">Salida: un archivo .xlsx listo para descarga.</div>
            </div>
            <div class="hero-card">
                <span>Plantilla base</span>
                <strong>FACTURACION REPUESTOS TYTSERV FEBRERO 2026.xlsx</strong>
                <div class="base-path"><?= htmlspecialchars($templateDir) ?></div>
            </div>
        </div>
    </section>

    <div class="layout">
        <aside class="panel stack">
            <h2>Entradas requeridas</h2>
            <?php foreach ($fileFields as $label): ?>
                <article class="box">
                    <span class="tag">Archivo</span>
                    <strong><?= htmlspecialchars($label) ?></strong>
                </article>
            <?php endforeach; ?>
            <article class="box">
                <span class="tag">Retencion</span>
                <strong><?= app_output_retention_limit() ?> ultimos reportes</strong>
                <div class="meta">Se elimina automaticamente lo mas antiguo.</div>
            </article>
        </aside>

        <main class="stack">
            <section class="panel">
                <h2>Procesar plantilla</h2>
                <p>Sube los 4 archivos del mes para generar el consolidado final.</p>

                <div class="form-grid">
                    <div>
                        <form method="post" enctype="multipart/form-data">
                            <?php foreach ($fileFields as $field => $label): ?>
                                <label class="file-label">
                                    <?= htmlspecialchars($label) ?>
                                    <input type="file" name="<?= htmlspecialchars($field) ?>" accept=".xls,.xlsx" required>
                                </label>
                            <?php endforeach; ?>
                            <button type="submit">Procesar y generar reporte</button>
                        </form>

                        <?php if ($error !== null): ?>
                            <div class="msg err"><?= htmlspecialchars($error) ?></div>
                        <?php endif; ?>

                        <?php if ($result !== null): ?>
                            <div class="msg ok">Archivo generado correctamente: <?= htmlspecialchars($result['excel_name']) ?></div>
                            <a class="button-link" href="<?= htmlspecialchars($result['download_url']) ?>">Descargar reporte final</a>
                            <?php if ($result['summary'] !== []): ?>
                                <div class="summary-grid" style="margin-top: 12px;">
                                    <?php foreach ($result['summary'] as $item): ?>
                                        <article class="box">
                                            <span class="tag"><?= htmlspecialchars((string)$item['label']) ?></span>
                                            <strong><?= (int)$item['rows'] ?> filas procesadas</strong>
                                        </article>
                                    <?php endforeach; ?>
                                </div>
                            <?php endif; ?>
                            <?php if ($result['console'] !== ''): ?>
                                <pre class="console"><?= htmlspecialchars($result['console']) ?></pre>
                            <?php endif; ?>
                        <?php endif; ?>
                    </div>

                    <article class="box">
                        <span class="tag">Salida</span>
                        <strong>storage/outputs</strong>
                        <div class="meta">Prefijo de archivos: repuestos_tytserv_</div>
                    </article>
                </div>
            </section>

            <section class="panel">
                <h2>Historial reciente</h2>
                <p>Se conservan solo los ultimos <?= app_output_retention_limit() ?> reportes de esta accion.</p>
                <div class="history" style="margin-top: 14px;">
                    <?php if ($history === []): ?>
                        <article class="box">
                            <strong>No hay reportes generados aun.</strong>
                            <div class="meta">La primera ejecucion aparecera aqui.</div>
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
