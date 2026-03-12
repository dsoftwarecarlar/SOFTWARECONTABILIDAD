<?php
declare(strict_types=1);

require dirname(__DIR__, 2) . '/includes/app.php';

$brand = app_brand();
$context = app_workspace_module('cxp', 'cxp_txt');
if ($context === null) {
    throw new RuntimeException('No se encontro la configuracion del modulo.');
}

$workspace = $context['workspace'];
$currentModule = $context['module'];
$uploadsDir = app_ensure_dir(app_storage_path('uploads'));
$outputsDir = app_ensure_dir(app_storage_path('outputs'));
$requestMethod = $_SERVER['REQUEST_METHOD'] ?? 'GET';

$result = null;
$error = null;

if ($requestMethod === 'POST') {
    try {
        if (!isset($_FILES['txt_file'])) {
            throw new RuntimeException('No se recibio el archivo TXT.');
        }

        $file = $_FILES['txt_file'];
        if (!is_array($file) || (int)($file['error'] ?? UPLOAD_ERR_NO_FILE) !== UPLOAD_ERR_OK) {
            throw new RuntimeException('Error al subir el archivo TXT.');
        }

        $originalName = (string)($file['name'] ?? 'archivo.txt');
        $tmpPath = (string)($file['tmp_name'] ?? '');
        $extension = strtolower(pathinfo($originalName, PATHINFO_EXTENSION));
        if ($extension !== 'txt') {
            throw new RuntimeException('Solo se permiten archivos TXT.');
        }

        $safeBase = preg_replace('/[^A-Za-z0-9_-]+/', '_', pathinfo($originalName, PATHINFO_FILENAME));
        $timestamp = date('Ymd_His');
        $inputFileName = $safeBase . '_' . $timestamp . '.txt';
        $outputFileName = $safeBase . '_' . $timestamp . '_accion2.xlsx';

        $inputPath = app_join_path($uploadsDir, $inputFileName);
        $outputPath = app_join_path($outputsDir, $outputFileName);

        if (!move_uploaded_file($tmpPath, $inputPath)) {
            throw new RuntimeException('No se pudo guardar el TXT subido.');
        }

        $scriptPath = app_join_path(app_root(), 'run_bot_accion2.js');
        if (!is_file($scriptPath)) {
            throw new RuntimeException('No existe run_bot_accion2.js en el proyecto.');
        }

        $command = escapeshellarg('node') . ' ' .
            escapeshellarg($scriptPath) . ' ' .
            escapeshellarg($inputPath) . ' ' .
            escapeshellarg($outputPath) . ' 2>&1';

        $lines = [];
        $exitCode = 0;
        exec($command, $lines, $exitCode);

        $consoleLines = [];
        foreach ($lines as $line) {
            $trimmed = trim((string)$line);
            if ($trimmed === '') {
                continue;
            }
            if (stripos($trimmed, 'Auditoria JSON:') === 0) {
                continue;
            }
            if (stripos($trimmed, 'TXT leido:') === 0) {
                continue;
            }
            if (stripos($trimmed, 'Excel generado (una sola hoja):') === 0) {
                continue;
            }
            $consoleLines[] = $trimmed;
        }
        $console = trim(implode(PHP_EOL, $consoleLines));

        if ($exitCode !== 0) {
            throw new RuntimeException($console === '' ? 'El proceso fallo.' : $console);
        }

        $generatedPath = $outputPath;
        foreach ($lines as $line) {
            $line = trim((string)$line);
            if (stripos($line, 'Excel generado (una sola hoja):') === 0) {
                $generatedPath = trim(substr($line, strlen('Excel generado (una sola hoja):')));
                break;
            }
        }

        $generatedReal = realpath($generatedPath) ?: $generatedPath;
        $generatedName = basename($generatedReal);
        if (!is_file($generatedReal)) {
            throw new RuntimeException('No se encontro el Excel generado.');
        }

        $result = [
            'excel_name' => $generatedName,
            'download_url' => app_output_download_url($generatedName),
            'console' => $console,
        ];

        app_cleanup_output_files_for_action('accion2', app_output_retention_limit());
    } catch (Throwable $e) {
        $error = $e->getMessage();
    }
}

$history = app_list_output_files_for_action('accion2', app_output_retention_limit());
$otherModules = array_values(array_filter(
    $workspace['modules'],
    static fn(array $module): bool => $module['slug'] !== 'cxp_txt'
));
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
            --accent-2: #bf7534;
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
        .page {
            max-width: 1280px;
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
            right: -70px;
            top: -76px;
            width: 280px;
            height: 280px;
            border-radius: 50%;
            background: radial-gradient(circle, rgba(15, 111, 103, 0.16), transparent 68%);
        }
        .hero-grid {
            position: relative;
            z-index: 1;
            display: grid;
            grid-template-columns: minmax(0, 1fr) 320px;
            gap: 20px;
            align-items: center;
        }
        h1 {
            margin: 14px 0 12px;
            font-family: "Bodoni MT", "Book Antiqua", serif;
            font-size: clamp(38px, 5.8vw, 70px);
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
            background: rgba(15, 111, 103, 0.09);
            color: var(--deep);
            font-weight: 700;
        }
        .hero-side {
            display: grid;
            gap: 12px;
        }
        .logo-stage {
            padding: 16px;
            border-radius: 28px;
            background: linear-gradient(180deg, rgba(22, 38, 48, 0.95), rgba(14, 28, 36, 0.96));
            box-shadow: inset 0 0 0 1px rgba(255, 255, 255, 0.07);
            display: grid;
            place-items: center;
        }
        .logo-frame {
            width: 176px;
            height: 176px;
            border-radius: 36px;
            display: grid;
            place-items: center;
            background: linear-gradient(180deg, rgba(255, 255, 255, 0.98), rgba(247, 238, 223, 0.95));
            box-shadow: 0 18px 34px rgba(0, 0, 0, 0.22);
        }
        .logo-frame img {
            width: 132px;
            height: 132px;
            object-fit: contain;
        }
        .hero-card {
            padding: 16px 18px;
            border-radius: 22px;
            background: rgba(255, 255, 255, 0.76);
            border: 1px solid rgba(23, 36, 45, 0.10);
        }
        .hero-card span {
            display: block;
            font-size: 12px;
            letter-spacing: 0.08em;
            text-transform: uppercase;
            color: var(--muted);
            font-weight: 700;
        }
        .hero-card strong {
            display: block;
            margin-top: 8px;
            font-size: 18px;
            line-height: 1.25;
        }
        .layout {
            margin-top: 18px;
            display: grid;
            grid-template-columns: 360px minmax(0, 1fr);
            gap: 16px;
            align-items: start;
        }
        .panel {
            border-radius: 28px;
            border: 1px solid var(--line);
            background: var(--panel);
            box-shadow: var(--shadow);
            padding: 22px;
        }
        .panel h2 {
            margin: 0 0 8px;
            font-family: "Franklin Gothic Medium", "Arial Narrow", sans-serif;
            font-size: 24px;
            letter-spacing: -0.03em;
        }
        .panel p {
            margin: 0;
            color: var(--muted);
            line-height: 1.58;
        }
        .side-stack,
        .main-stack,
        .history,
        .module-links,
        .stats {
            display: grid;
            gap: 12px;
        }
        .map-item,
        .module-link,
        .stats-item,
        .history-item {
            padding: 16px 18px;
            border-radius: 20px;
            border: 1px solid rgba(23, 36, 45, 0.10);
            background: rgba(255, 255, 255, 0.78);
        }
        .map-item span,
        .stats-item span {
            display: block;
            font-size: 12px;
            letter-spacing: 0.08em;
            text-transform: uppercase;
            color: var(--muted);
            font-weight: 700;
        }
        .map-item strong,
        .stats-item strong,
        .module-link strong {
            display: block;
            margin-top: 8px;
            font-size: 18px;
            line-height: 1.25;
        }
        .map-item p,
        .stats-item p,
        .module-link p {
            margin-top: 8px;
            color: var(--muted);
            font-size: 14px;
            line-height: 1.5;
        }
        .module-link {
            text-decoration: none;
            color: inherit;
        }
        .module-link:hover {
            border-color: rgba(15, 111, 103, 0.25);
        }
        .form-grid {
            display: grid;
            grid-template-columns: minmax(0, 1fr) 300px;
            gap: 16px;
            margin-top: 14px;
            align-items: start;
        }
        form {
            display: grid;
            gap: 12px;
        }
        input[type="file"] {
            width: 100%;
            padding: 18px;
            border-radius: 18px;
            border: 1px dashed rgba(15, 111, 103, 0.34);
            background: rgba(255, 255, 255, 0.75);
        }
        button, .button-link {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            padding: 12px 18px;
            border-radius: 999px;
            border: 0;
            background: linear-gradient(135deg, var(--accent) 0%, var(--deep) 100%);
            color: #fff;
            text-decoration: none;
            font-weight: 700;
            cursor: pointer;
        }
        .msg {
            margin-top: 14px;
            padding: 14px 16px;
            border-radius: 18px;
            white-space: pre-wrap;
        }
        .msg.ok {
            background: var(--ok-bg);
            color: var(--ok-ink);
        }
        .msg.err {
            background: var(--err-bg);
            color: var(--err-ink);
        }
        .console {
            margin-top: 12px;
            padding: 14px;
            border-radius: 18px;
            background: #18242c;
            color: #dfe7e4;
            font-family: Consolas, "Courier New", monospace;
            font-size: 12px;
            max-height: 240px;
            overflow: auto;
        }
        .history-item {
            display: grid;
            grid-template-columns: 1fr auto;
            align-items: center;
            gap: 12px;
        }
        .history-item strong {
            display: block;
        }
        .history-meta {
            margin-top: 6px;
            color: var(--muted);
            font-size: 13px;
        }
        @media (max-width: 980px) {
            .hero-grid,
            .layout,
            .form-grid,
            .history-item {
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
            <strong><?= htmlspecialchars($currentModule['title']) ?></strong>
            <span class="note">Accion 2 del flujo contable mensual</span>
        </div>
        <a class="back-link" href="<?= htmlspecialchars(app_url('modules/cxp_pdf/index.php')) ?>">Volver a Accion 1</a>
    </div>

    <section class="hero">
        <div class="hero-grid">
            <div>
                <span class="chip">Flujo TXT en produccion</span>
                <h1>TXT desordenado a Excel RET PROV con formato exacto del ejemplo manual.</h1>
                <p class="lead">
                    Este modulo toma un TXT de retenciones, normaliza filas y columnas, y genera un Excel final
                    en hoja RET PROV siguiendo el formato de `ACCION2.xlsx` dentro de `EJEMPLOSAMANO`.
                </p>
                <div class="hero-note">Salida de usuario: descarga de archivo .xlsx.</div>
            </div>
            <div class="hero-side">
                <?php if ($brand['logo_exists']): ?>
                    <div class="logo-stage">
                        <div class="logo-frame">
                            <img src="<?= htmlspecialchars($brand['logo_url']) ?>" alt="<?= htmlspecialchars($brand['company']) ?>">
                        </div>
                    </div>
                <?php endif; ?>
                <div class="hero-card">
                    <span>Plantilla base</span>
                    <strong>outputs/EJEMPLOSAMANO/ACCION2.xlsx</strong>
                </div>
            </div>
        </div>
    </section>

    <div class="layout">
        <aside class="side-stack">
            <section class="panel">
                <h2>Mapa del proceso</h2>
                <div class="map-item">
                    <span>Entrada</span>
                    <strong>Archivo .txt</strong>
                    <p>Cargado por el usuario del area para la accion mensual.</p>
                </div>
                <div class="map-item">
                    <span>Transformacion</span>
                    <strong>Normalizacion y clasificacion</strong>
                    <p>Se validan columnas clave: NUM RT, FECHA, TIPO, BASE, RETENCION.</p>
                </div>
                <div class="map-item">
                    <span>Salida</span>
                    <strong>Excel RET PROV</strong>
                    <p>Archivo final con formato de ejemplo manual y resumen lateral.</p>
                </div>
            </section>

            <?php if (count($otherModules) > 0): ?>
                <section class="panel">
                    <h2>Otras acciones del workspace</h2>
                    <div class="module-links">
                        <?php foreach ($otherModules as $module): ?>
                            <a class="module-link" href="<?= htmlspecialchars((string)$module['url']) ?>">
                                <strong><?= htmlspecialchars($module['title']) ?></strong>
                                <p><?= htmlspecialchars($module['description']) ?></p>
                            </a>
                        <?php endforeach; ?>
                    </div>
                </section>
            <?php endif; ?>
        </aside>

        <main class="main-stack">
            <section class="panel">
                <h2>Procesar archivo TXT</h2>
                <p>Sube el TXT y el sistema devolvera el Excel final para esta segunda accion.</p>

                <div class="form-grid">
                    <div>
                        <form method="post" enctype="multipart/form-data">
                            <input type="file" name="txt_file" accept=".txt,text/plain" required>
                            <button type="submit">Generar Excel</button>
                        </form>

                        <?php if ($error !== null): ?>
                            <div class="msg err"><?= htmlspecialchars($error) ?></div>
                        <?php endif; ?>

                        <?php if ($result !== null): ?>
                            <div class="msg ok">Archivo generado correctamente: <?= htmlspecialchars($result['excel_name']) ?></div>
                            <a class="button-link" href="<?= htmlspecialchars($result['download_url']) ?>">Descargar Excel</a>
                            <?php if ($result['console'] !== ''): ?>
                                <pre class="console"><?= htmlspecialchars($result['console']) ?></pre>
                            <?php endif; ?>
                        <?php endif; ?>
                    </div>

                    <div class="stats">
                        <div class="stats-item">
                            <span>Salida publica</span>
                            <strong>Solo .xlsx</strong>
                            <p>No se expone el JSON interno en la interfaz.</p>
                        </div>
                        <div class="stats-item">
                            <span>Historial</span>
                            <strong>Ordenado por creacion</strong>
                            <p>Ultimos archivos arriba, mas antiguos abajo.</p>
                        </div>
                        <div class="stats-item">
                            <span>Destino de archivos</span>
                            <strong>storage/outputs</strong>
                            <p>Ubicacion central para control del equipo.</p>
                        </div>
                    </div>
                </div>
            </section>

            <section class="panel">
                <h2>Historial reciente (Accion 2)</h2>
                <p>Archivos generados por este modulo TXT a Excel. Se conservan solo los ultimos <?= app_output_retention_limit() ?> informes.</p>
                <div class="history" style="margin-top: 14px;">
                    <?php if (count($history) === 0): ?>
                        <div class="history-item">
                            <div>
                                <strong>Sin archivos de Accion 2 todavia</strong>
                                <div class="history-meta">Se mostraran aqui cuando ejecutes el primer TXT.</div>
                            </div>
                        </div>
                    <?php else: ?>
                        <?php foreach ($history as $item): ?>
                            <div class="history-item">
                                <div>
                                    <strong><?= htmlspecialchars($item['name']) ?></strong>
                                    <div class="history-meta"><?= htmlspecialchars($item['time']) ?> | <?= number_format((int)$item['size'] / 1024, 2) ?> KB</div>
                                </div>
                                <div>
                                    <a class="button-link" href="<?= htmlspecialchars(app_output_download_url($item['name'])) ?>">Descargar</a>
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
