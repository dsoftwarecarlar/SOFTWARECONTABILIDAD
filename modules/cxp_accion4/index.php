<?php
declare(strict_types=1);

require dirname(__DIR__, 2) . '/includes/app.php';

$brand = app_brand();
$context = app_workspace_module('cxp', 'cxp_accion4');
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

        $originalName = trim((string)($file['name'] ?? 'documento.txt'));
        if ($originalName === '') {
            $originalName = 'documento.txt';
        }

        $timestamp = date('Ymd_His');
        $tmpPath = (string)($file['tmp_name'] ?? '');
        $extension = strtolower(pathinfo($originalName, PATHINFO_EXTENSION));
        if ($extension !== 'txt') {
            throw new RuntimeException('Solo se permiten archivos TXT.');
        }

        $baseName = preg_replace('/[^A-Za-z0-9_-]+/', '_', pathinfo($originalName, PATHINFO_FILENAME));
        $baseName = trim((string)$baseName, '_');
        if ($baseName === '') {
            $baseName = 'documento';
        }

        $inputFileName = $baseName . '_' . $timestamp . '.txt';
        $inputPath = app_join_path($uploadsDir, $inputFileName);
        if (!move_uploaded_file($tmpPath, $inputPath)) {
            throw new RuntimeException('No se pudo guardar el TXT subido.');
        }
        $inputPaths = [$inputPath];

        $outputFileName = $baseName . '_' . $timestamp . '_accion4.xlsx';
        $outputPath = app_join_path($outputsDir, $outputFileName);

        $scriptPath = app_join_path(app_root(), 'run_bot_accion4.js');
        if (!is_file($scriptPath)) {
            throw new RuntimeException('No existe run_bot_accion4.js en el proyecto.');
        }

        $commandParts = [
            escapeshellarg('node'),
            escapeshellarg($scriptPath),
        ];

        $commandParts[] = escapeshellarg($inputPath);
        $commandParts[] = escapeshellarg('--output');
        $commandParts[] = escapeshellarg($outputPath);

        $command = implode(' ', $commandParts) . ' 2>&1';

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
            if (stripos($trimmed, 'Archivo leido:') === 0) {
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

        app_cleanup_output_files_for_action('accion4', app_output_retention_limit());
    } catch (Throwable $e) {
        $error = $e->getMessage();
    }
}

$latestActions = app_latest_action_exports();
$missingActions = array_values(array_filter(
    $latestActions,
    static fn(array $item): bool => $item['latest'] === null
));
$canExportAllActions = count($missingActions) === 0;
$history = app_list_output_files_for_action('accion4', app_output_retention_limit());
?>
<!doctype html>
<html lang="es">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title><?= htmlspecialchars($workspace['title']) ?> | <?= htmlspecialchars($currentModule['title']) ?></title>
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
        .hero-side {
            display: grid;
            gap: 12px;
        }
        .logo-stage {
            padding: 16px;
            border-radius: 28px;
            background: linear-gradient(180deg, rgba(21, 38, 48, 0.95), rgba(14, 28, 37, 0.96));
            box-shadow: inset 0 0 0 1px rgba(255, 255, 255, 0.07);
            display: grid;
            place-items: center;
        }
        .logo-frame {
            width: 180px;
            height: 180px;
            border-radius: 38px;
            display: grid;
            place-items: center;
            background: linear-gradient(180deg, rgba(255, 255, 255, 0.98), rgba(248, 239, 224, 0.95));
            box-shadow: 0 18px 34px rgba(0, 0, 0, 0.22);
        }
        .logo-frame img {
            width: 136px;
            height: 136px;
            object-fit: contain;
        }
        .hero-card {
            padding: 16px 18px;
            border-radius: 22px;
            background: rgba(255, 255, 255, 0.76);
            border: 1px solid rgba(24, 36, 44, 0.10);
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
            font-size: 25px;
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
        .stats {
            display: grid;
            gap: 12px;
        }
        .map-item,
        .stats-item,
        .history-item {
            padding: 16px 18px;
            border-radius: 20px;
            border: 1px solid rgba(24, 36, 44, 0.10);
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
        .stats-item strong {
            display: block;
            margin-top: 8px;
            font-size: 19px;
            line-height: 1.25;
        }
        .map-item p,
        .stats-item p {
            margin-top: 8px;
            color: var(--muted);
            font-size: 14px;
            line-height: 1.5;
        }
        .map-item a {
            display: inline-flex;
            align-items: center;
            margin-top: 10px;
            color: var(--accent);
            font-weight: 700;
            text-decoration: none;
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
            border: 1px dashed rgba(16, 111, 102, 0.34);
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
            <strong><?= htmlspecialchars($workspace['title']) ?></strong>
            <span class="note">Modulo operativo para procesamiento contable en produccion</span>
        </div>
        <a class="back-link" href="<?= htmlspecialchars(app_url()) ?>">Volver al portal</a>
    </div>

    <section class="hero">
        <div class="hero-grid">
            <div>
                <span class="chip">Flujo productivo activo</span>
                <h1><?= htmlspecialchars($currentModule['title']) ?></h1>
                <p class="lead">
                    Flujo en produccion: carga de un solo TXT MAYOR GENERAL IVA y generacion de Excel MAYOR IVA
                    respetando exactamente la estructura del ejemplo manual en hoja unica.
                </p>
                <div class="hero-note">Salida de usuario: descarga directa de archivo .xlsx.</div>
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
                    <span>Alcance actual</span>
                    <strong>Proceso principal de Cuentas por Pagar en operacion diaria.</strong>
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
                    <strong>Un TXT MAYOR GENERAL IVA</strong>
                    <p>Documento fuente tipo MAYOR GENERAL IVA para devolver un solo Excel igual al ejemplo manual.</p>
                </div>
                <div class="map-item">
                    <span>Transformacion</span>
                    <strong>Replica exacta de estructura manual</strong>
                    <p>El sistema omite, reordena e inserta filas de formula para igualar el formato operativo esperado.</p>
                </div>
                <div class="map-item">
                    <span>Salida</span>
                    <strong>Excel final MAYOR IVA</strong>
                    <p>Archivo listo para descarga y uso operativo del departamento.</p>
                </div>
                <div class="map-item">
                    <span>Control interno</span>
                    <strong>Auditoria tecnica en storage</strong>
                    <p>Registro interno del proceso para revision del equipo tecnico.</p>
                </div>
                <div class="map-item">
                    <span>Siguiente accion</span>
                    <strong>Clasificacion TXT a Excel</strong>
                    <p>Proceso principal de proveedores (accion 1).</p>
                    <a href="<?= htmlspecialchars(app_url('modules/cxp_pdf/index.php')) ?>">Abrir Accion 1</a>
                </div>
            </section>
        </aside>

        <main class="main-stack">
            <section class="panel">
                <h2>Procesar TXT MAYOR IVA</h2>
                <p>Sube un solo TXT y el sistema devolvera un Excel MAYOR IVA igual al ejemplo manual cargado en el proyecto.</p>

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
                            <p>La interfaz publica no expone el JSON de auditoria.</p>
                        </div>
                        <div class="stats-item">
                            <span>Ubicacion de archivos</span>
                            <strong>storage/outputs</strong>
                            <p>Centralizado para respaldo y control del equipo.</p>
                        </div>
                        <div class="stats-item">
                            <span>Disponibilidad</span>
                            <strong>Operacion interna</strong>
                            <p>Disenado para uso local del departamento contable.</p>
                        </div>
                        <div class="stats-item">
                            <span>Consolidado total</span>
                            <strong><?= $canExportAllActions ? 'Listo para descargar' : 'Pendiente de completar' ?></strong>
                            <?php if ($canExportAllActions): ?>
                                <p>Genera un solo Excel con el ultimo archivo producido en cada accion.</p>
                                <a class="button-link" href="<?= htmlspecialchars(app_url('export_all_actions.php')) ?>">Descargar todas las acciones</a>
                            <?php else: ?>
                                <p>Faltan salidas recientes para: <?= htmlspecialchars(implode(', ', array_map(
                                    static fn(array $item): string => (string)$item['label'],
                                    $missingActions
                                ))) ?>.</p>
                            <?php endif; ?>
                        </div>
                    </div>
                </div>
            </section>

            <section class="panel">
                <h2>Ultimo Excel por accion</h2>
                <p>Base usada para armar el consolidado general.</p>
                <div class="history" style="margin-top: 14px;">
                    <?php foreach ($latestActions as $action): ?>
                        <div class="history-item">
                            <div>
                                <strong><?= htmlspecialchars($action['label']) ?></strong>
                                <?php if ($action['latest'] !== null): ?>
                                    <div class="history-meta">
                                        <?= htmlspecialchars((string)$action['latest']['name']) ?>
                                        | <?= htmlspecialchars((string)$action['latest']['time']) ?>
                                    </div>
                                <?php else: ?>
                                    <div class="history-meta">Todavia no hay Excel generado para esta accion.</div>
                                <?php endif; ?>
                            </div>
                            <div>
                                <?php if ($action['latest'] !== null): ?>
                                    <a class="button-link" href="<?= htmlspecialchars(app_output_download_url((string)$action['latest']['name'])) ?>">Descargar</a>
                                <?php else: ?>
                                    <a class="button-link" href="<?= htmlspecialchars((string)$action['module_url']) ?>">Abrir accion</a>
                                <?php endif; ?>
                            </div>
                        </div>
                    <?php endforeach; ?>
                </div>
            </section>

            <section class="panel">
                <h2>Historial reciente</h2>
                <p>Archivos generados por este modulo de produccion. Se conservan solo los ultimos <?= app_output_retention_limit() ?> informes.</p>
                <div class="history" style="margin-top: 14px;">
                    <?php if (count($history) === 0): ?>
                        <div class="history-item">
                            <div>
                                <strong>Sin archivos generados todavia</strong>
                                <div class="history-meta">El historial se llena con cada ejecucion del proceso.</div>
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
