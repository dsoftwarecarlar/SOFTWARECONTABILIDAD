<?php
declare(strict_types=1);

require dirname(__DIR__, 2) . '/includes/app.php';
require dirname(__DIR__, 2) . '/includes/servicios_marcas_job_runner.php';

$brand = app_brand();
$context = app_workspace_module('cxp', 'cxp_servicios_marcas');
if ($context === null) {
    throw new RuntimeException('No se encontro la configuracion del modulo.');
}

$windowContext = app_workspace_window('cxp', 'conciliacion_servicios_marcas');
$workspace = $context['workspace'];
$currentModule = $context['module'];
$window = $windowContext['window'] ?? null;
$uploadsDir = app_ensure_dir(app_storage_path('uploads'));
$outputsDir = app_ensure_dir(app_storage_path('outputs'));
$templateDir = app_join_path(app_root(), 'outputs', 'EJEMPLOAMANOTAREA2');
$jobsDir = app_ensure_dir(app_storage_path('jobs'));
$requestMethod = $_SERVER['REQUEST_METHOD'] ?? 'GET';

$result = null;
$error = null;
$notice = null;
$noticeConsole = '';
$activeJobId = '';
$activeJob = null;
$pendingJob = null;

$historyLabel = static function (string $fileName): string {
    $name = strtolower($fileName);
    return match (true) {
        str_starts_with($name, 'servicios_changan_') => 'CHANGAN',
        str_starts_with($name, 'servicios_peug_') => 'PEUGEOT',
        str_starts_with($name, 'servicios_szk_') => 'SUZUKI',
        str_starts_with($name, 'servicios_tyt_') => 'MATRIZ',
        default => 'SERVICIOS',
    };
};

servicios_job_refresh_stale_jobs($jobsDir);

$readJob = static function (string $jobId) use ($jobsDir): ?array {
    if ($jobId === '' || preg_match('/^[A-Za-z0-9_-]+$/', $jobId) !== 1) {
        return null;
    }

    $path = app_join_path($jobsDir, 'servicios_marcas_' . $jobId . '.json');
    if (!is_file($path)) {
        return null;
    }

    $raw = file_get_contents($path);
    if ($raw === false || trim($raw) === '') {
        return null;
    }

    $data = json_decode($raw, true);
    return is_array($data) ? $data : null;
};

$writeJob = static function (string $jobId, array $payload) use ($jobsDir): void {
    $path = app_join_path($jobsDir, 'servicios_marcas_' . $jobId . '.json');
    file_put_contents($path, json_encode($payload, JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES), LOCK_EX);
};

$jobEntries = servicios_job_entries($jobsDir);
$activeJobs = array_values(array_filter(
    $jobEntries,
    static fn(array $job): bool => in_array((string)($job['status'] ?? ''), servicios_job_active_statuses(), true)
));

$stoppedCount = filter_input(INPUT_GET, 'stopped', FILTER_VALIDATE_INT);
if ($stoppedCount !== null && $stoppedCount !== false) {
    $notice = $stoppedCount > 0
        ? 'Se solicito detener ' . $stoppedCount . ' proceso(s) activo(s).'
        : 'No habia procesos activos para detener.';
}

$statusJobId = trim((string)($_GET['status'] ?? ''));
if ($statusJobId !== '') {
    header('Content-Type: application/json; charset=utf-8');

    $job = $readJob($statusJobId);
    if ($job === null) {
        http_response_code(404);
        echo json_encode(['status' => 'missing', 'message' => 'No se encontro el trabajo solicitado.'], JSON_UNESCAPED_SLASHES);
        exit;
    }

    echo json_encode($job, JSON_UNESCAPED_SLASHES);
    exit;
}

$activeJobId = trim((string)($_GET['job'] ?? ''));
if ($activeJobId !== '') {
    $activeJob = $readJob($activeJobId);
    if ($activeJob === null) {
        $error = 'No se encontro el proceso solicitado.';
        $activeJobId = '';
    } else {
        $jobStatus = (string)($activeJob['status'] ?? '');
        if ($jobStatus === 'complete') {
            $result = [
                'source_name' => (string)($activeJob['source_name'] ?? ''),
                'downloads' => is_array($activeJob['downloads'] ?? null) ? $activeJob['downloads'] : [],
                'summary' => is_array($activeJob['summary'] ?? null) ? $activeJob['summary'] : [],
                'console' => (string)($activeJob['console'] ?? ''),
            ];
        } elseif (in_array($jobStatus, ['queued', 'running', 'cancel_requested'], true)) {
            $pendingJob = $activeJob;
        } elseif ($jobStatus === 'cancelled') {
            $notice = (string)($activeJob['message'] ?? 'Proceso cancelado.');
            $noticeConsole = (string)($activeJob['console'] ?? '');
        } elseif ($jobStatus === 'error') {
            $error = (string)($activeJob['error'] ?? 'El proceso fallo.');
        }
    }
}

if ($requestMethod === 'POST') {
    try {
        $action = trim((string)($_POST['action'] ?? 'process'));
        if ($action === 'stop_all') {
            $stopResult = servicios_job_request_cancel_all($jobsDir);
            $returnJobId = trim((string)($_POST['return_job'] ?? ''));
            $redirectParams = ['stopped' => (string)$stopResult['count']];
            if ($returnJobId !== '') {
                $redirectParams['job'] = $returnJobId;
            }

            header('Location: ' . app_url('modules/cxp_servicios_marcas/index.php?' . http_build_query($redirectParams)));
            exit;
        }

        ignore_user_abort(true);
        set_time_limit(0);

        if (!is_dir($templateDir)) {
            throw new RuntimeException('No existe la carpeta base de plantillas mensuales.');
        }

        if (!isset($_FILES['excel_file'])) {
            throw new RuntimeException('No se recibio el archivo Excel.');
        }

        $file = $_FILES['excel_file'];
        if (!is_array($file) || (int)($file['error'] ?? UPLOAD_ERR_NO_FILE) !== UPLOAD_ERR_OK) {
            throw new RuntimeException('Error al subir el archivo Excel.');
        }

        $originalName = trim((string)($file['name'] ?? 'reporte.xls'));
        $tmpPath = (string)($file['tmp_name'] ?? '');
        $extension = strtolower(pathinfo($originalName, PATHINFO_EXTENSION));
        if (!in_array($extension, ['xls', 'xlsx'], true)) {
            throw new RuntimeException('Solo se permiten archivos Excel .xls o .xlsx.');
        }

        $safeBase = preg_replace('/[^A-Za-z0-9_-]+/', '_', pathinfo($originalName, PATHINFO_FILENAME));
        $safeBase = trim((string)$safeBase, '_');
        if ($safeBase === '') {
            $safeBase = 'reporte_servicios';
        }

        $timestamp = date('Ymd_His');
        $inputFileName = sprintf('%s_%s.%s', $safeBase, $timestamp, $extension);
        $inputPath = app_join_path($uploadsDir, $inputFileName);

        if (!move_uploaded_file($tmpPath, $inputPath)) {
            throw new RuntimeException('No se pudo guardar el Excel subido.');
        }

        $activeJobId = 'servicios_' . date('Ymd_His') . '_' . bin2hex(random_bytes(4));
        $writeJob($activeJobId, [
            'job_id' => $activeJobId,
            'status' => 'queued',
            'source_name' => $originalName,
            'message' => 'Proceso en cola. La pagina se actualizara automaticamente.',
            'created_at' => date('Y-m-d H:i:s'),
        ]);

        $redirectUrl = app_url('modules/cxp_servicios_marcas/index.php?job=' . rawurlencode($activeJobId));
        header('Location: ' . $redirectUrl);
        header('Content-Length: 0');
        header('Connection: close');
        if (function_exists('session_write_close')) {
            session_write_close();
        }
        while (ob_get_level() > 0) {
            @ob_end_flush();
        }
        flush();

        try {
            servicios_job_run($activeJobId, $inputPath, $outputsDir, $templateDir);
        } catch (Throwable $workerError) {
            // El estado ya queda registrado en el JSON del job para el polling del cliente.
        }
        exit;
    } catch (Throwable $e) {
        $error = $e->getMessage();
    }
}

$history = [];
$seenHistoryBrands = [];
foreach (app_list_output_files_for_action('servicios', 20) as $item) {
    $name = strtolower((string)($item['name'] ?? ''));
    $brandKey = null;
    foreach (servicios_job_output_config() as $key => $config) {
        if (str_starts_with($name, strtolower((string)$config['prefix']))) {
            $brandKey = $key;
            break;
        }
    }

    if ($brandKey === null || isset($seenHistoryBrands[$brandKey])) {
        continue;
    }

    $seenHistoryBrands[$brandKey] = true;
    $history[] = $item;
    if (count($history) >= 4) {
        break;
    }
}
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
        .topbar, .history-item, .download-card, .control-row { display: flex; justify-content: space-between; gap: 14px; align-items: center; }
        .topbar { margin-bottom: 16px; }
        .chip { display: inline-flex; padding: 6px 10px; border-radius: 999px; background: rgba(15,111,103,.1); color: var(--accent); font-size: 12px; font-weight: 700; letter-spacing: .08em; text-transform: uppercase; }
        .topbar strong { display: block; margin-top: 8px; font-family: "Franklin Gothic Medium", "Arial Narrow", sans-serif; font-size: 28px; letter-spacing: -.03em; }
        .topbar span.note { display: block; margin-top: 3px; color: var(--muted); font-size: 14px; }
        .back-link, .link-btn { color: var(--deep); text-decoration: none; font-weight: 700; }
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
        .hero-side { display: grid; gap: 12px; }
        .logo-stage, .hero-card { padding: 16px; border-radius: 24px; }
        .logo-stage { background: linear-gradient(180deg, rgba(21,38,48,.95), rgba(14,28,37,.96)); display: grid; place-items: center; }
        .logo-frame { width: 180px; height: 180px; border-radius: 38px; display: grid; place-items: center; background: linear-gradient(180deg, rgba(255,255,255,.98), rgba(248,239,224,.95)); }
        .logo-frame img { width: 136px; height: 136px; object-fit: contain; }
        .hero-card { background: rgba(255,255,255,.76); border: 1px solid rgba(24,36,44,.1); }
        .hero-card span, .tag { display: block; font-size: 12px; color: var(--muted); font-weight: 700; letter-spacing: .08em; text-transform: uppercase; }
        .hero-card strong, .box strong { display: block; margin-top: 8px; font-size: 18px; line-height: 1.3; }
        .layout { margin-top: 18px; display: grid; grid-template-columns: 360px minmax(0,1fr); gap: 16px; align-items: start; }
        .panel { background: var(--panel); padding: 22px; }
        .panel h2 { margin: 0 0 8px; font-family: "Franklin Gothic Medium", "Arial Narrow", sans-serif; font-size: 25px; letter-spacing: -.03em; }
        .stack, .brands, .history, .downloads, .summary-grid { display: grid; gap: 12px; }
        .box { padding: 16px 18px; border-radius: 20px; border: 1px solid rgba(24,36,44,.1); background: rgba(255,255,255,.78); }
        .box p { margin-top: 8px; font-size: 14px; }
        .form-grid { display: grid; grid-template-columns: minmax(0,1fr) 280px; gap: 16px; margin-top: 14px; align-items: start; }
        form { display: grid; gap: 12px; }
        input[type="file"] { width: 100%; padding: 18px; border-radius: 18px; border: 1px dashed rgba(16,111,102,.34); background: rgba(255,255,255,.75); }
        button, .button-link { display: inline-flex; align-items: center; justify-content: center; padding: 12px 18px; border-radius: 999px; border: 0; background: linear-gradient(135deg, var(--accent) 0%, var(--deep) 100%); color: #fff; text-decoration: none; font-weight: 700; cursor: pointer; }
        button[disabled] { opacity: .7; cursor: progress; }
        .stop-button { background: linear-gradient(135deg, #9d2f28 0%, #6f1c18 100%); }
        .msg { margin-top: 14px; padding: 14px 16px; border-radius: 18px; white-space: pre-wrap; }
        .ok { background: var(--ok-bg); color: var(--ok-ink); }
        .err { background: var(--err-bg); color: var(--err-ink); }
        .pending { background: rgba(227, 239, 238, 0.95); color: var(--deep); border: 1px solid rgba(15, 111, 103, 0.18); }
        .console { margin-top: 12px; padding: 14px; border-radius: 18px; background: #18242c; color: #dfe7e4; font-family: Consolas, "Courier New", monospace; font-size: 12px; max-height: 260px; overflow: auto; }
        .meta { margin-top: 6px; color: var(--muted); font-size: 13px; }
        @media (max-width: 980px) {
            .hero-grid, .layout, .form-grid { grid-template-columns: 1fr; }
            .topbar, .history-item, .download-card, .control-row { align-items: flex-start; flex-direction: column; }
        }
    </style>
</head>
<body>
<div class="page">
    <div class="topbar">
        <div>
            <span class="chip"><?= htmlspecialchars($brand['platform']) ?></span>
            <strong><?= htmlspecialchars($workspace['title']) ?></strong>
            <span class="note">Modulo operativo para separar el reporte mensual por marca</span>
        </div>
        <a class="back-link" href="<?= htmlspecialchars((string)($window['url'] ?? ($workspace['home_url'] ?? app_url()))) ?>">Volver a la ventana</a>
    </div>

    <section class="hero">
        <div class="hero-grid">
            <div>
                <span class="chip">Flujo productivo activo</span>
                <h1><?= htmlspecialchars($currentModule['title']) ?></h1>
                <p class="lead">Carga un Excel mensual con todas las marcas, separa los registros y genera una salida final independiente por plantilla.</p>
                <div class="hero-note">Salida de usuario: descarga directa de archivos .xls por marca.</div>
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
                    <span>Acceso LAN</span>
                    <strong><?= htmlspecialchars(app_public_root_url()) ?></strong>
                </div>
            </div>
        </div>
    </section>

    <div class="layout">
        <aside class="panel">
            <h2>Plantillas</h2>
            <p>Las 4 salidas se generan usando la base mensual del area y se conservan solo los 3 ultimos archivos por marca.</p>
            <div class="brands" style="margin-top: 14px;">
                <article class="box"><span class="tag">CHANGAN</span><strong>Plantilla CHANGAN</strong><p>Datos filtrados por marca y escritos en su libro operativo.</p></article>
                <article class="box"><span class="tag">PEUGEOT</span><strong>Plantilla PEUGEOT</strong><p>Facturas y notas quedan separadas en la base visible del mes.</p></article>
                <article class="box"><span class="tag">SUZUKI</span><strong>Plantilla SUZUKI</strong><p>Consolida SUZUKI AMBATO y SUZUKI RIOBAMBA en la plantilla correspondiente.</p></article>
                <article class="box"><span class="tag">MATRIZ</span><strong>Plantilla MATRIZ</strong><p>Se respeta el formato mensual existente para la marca MATRIZ.</p></article>
                <article class="box"><span class="tag">Base mensual</span><strong><?= htmlspecialchars($templateDir) ?></strong><p>La carpeta contiene las plantillas del mes y el ejemplo operativo de referencia.</p></article>
            </div>
        </aside>

        <main class="stack">
            <section class="panel">
                <h2>Cargar Excel</h2>
                <p>Sube el reporte mensual descargado del sistema. El proceso acepta .xls o .xlsx y devuelve una descarga separada por marca.</p>
                <div class="form-grid">
                    <form method="post" enctype="multipart/form-data" data-job-form>
                        <input type="hidden" name="action" value="process">
                        <input type="file" name="excel_file" accept=".xls,.xlsx" required>
                        <button type="submit" data-submit-button>Procesar y generar plantillas</button>
                    </form>
                    <article class="box">
                        <span class="tag">Retencion</span>
                        <strong>3 archivos por marca</strong>
                        <p>La limpieza automatica evita acumulacion en `storage/outputs`.</p>
                    </article>
                </div>
                <div class="control-row" style="margin-top: 12px;">
                    <form method="post" data-stop-form>
                        <input type="hidden" name="action" value="stop_all">
                        <?php if ($activeJobId !== ''): ?>
                            <input type="hidden" name="return_job" value="<?= htmlspecialchars($activeJobId) ?>">
                        <?php endif; ?>
                        <button type="submit" class="stop-button" data-stop-button <?= $activeJobs === [] ? 'disabled' : '' ?>>Parar todos los procesos</button>
                    </form>
                    <div class="meta">Procesos activos: <?= count($activeJobs) ?> - Historial visible: ultimo archivo por marca.</div>
                </div>

                <?php if ($pendingJob !== null): ?>
                    <div
                        class="msg pending"
                        data-job-box
                        data-job-id="<?= htmlspecialchars($activeJobId) ?>"
                        data-status-url="<?= htmlspecialchars(app_url('modules/cxp_servicios_marcas/index.php?status=' . rawurlencode($activeJobId))) ?>"
                        data-return-url="<?= htmlspecialchars(app_url('modules/cxp_servicios_marcas/index.php?job=' . rawurlencode($activeJobId))) ?>"
                    >
                        <strong data-job-message><?= htmlspecialchars((string)($pendingJob['message'] ?? 'Procesando en segundo plano.')) ?></strong>
                        <div class="meta" data-job-meta>
                            Estado: <?= htmlspecialchars((string)($pendingJob['status'] ?? 'running')) ?>
                            <?php if (!empty($pendingJob['started_at'])): ?>
                                · Inicio: <?= htmlspecialchars((string)$pendingJob['started_at']) ?>
                            <?php elseif (!empty($pendingJob['created_at'])): ?>
                                · Creado: <?= htmlspecialchars((string)$pendingJob['created_at']) ?>
                            <?php endif; ?>
                        </div>
                        <div class="meta">Puedes dejar esta pagina abierta. Se actualizara sola cuando termine.</div>
                    </div>
                <?php endif; ?>

                <?php if ($result !== null): ?>
                    <div class="msg ok">Archivo procesado: <?= htmlspecialchars($result['source_name']) ?>.</div>
                    <?php if ($result['summary'] !== []): ?>
                        <div class="summary-grid" style="margin-top: 12px;">
                            <?php foreach ($result['summary'] as $item): ?>
                                <article class="box">
                                    <span class="tag"><?= htmlspecialchars((string)($item['label'] ?? 'MARCA')) ?></span>
                                    <strong><?= (int)($item['rows'] ?? 0) ?> filas procesadas</strong>
                                    <div class="meta">
                                        Facturas fallback: <?= (int)($item['invoice_fallbacks'] ?? 0) ?>
                                        · Notas fallback: <?= (int)($item['note_fallbacks'] ?? 0) ?>
                                    </div>
                                </article>
                            <?php endforeach; ?>
                        </div>
                    <?php endif; ?>
                    <div class="downloads" style="margin-top: 12px;">
                        <?php foreach ($result['downloads'] as $download): ?>
                            <article class="box download-card">
                                <div>
                                    <span class="tag"><?= htmlspecialchars($download['label']) ?></span>
                                    <strong><?= htmlspecialchars($download['name']) ?></strong>
                                </div>
                                <a class="button-link" href="<?= htmlspecialchars($download['download_url']) ?>">Descargar</a>
                            </article>
                        <?php endforeach; ?>
                    </div>
                    <?php if ($result['console'] !== ''): ?>
                        <pre class="console"><?= htmlspecialchars($result['console']) ?></pre>
                    <?php endif; ?>
                <?php endif; ?>

                <?php if ($notice !== null): ?>
                    <div class="msg pending"><?= htmlspecialchars($notice) ?></div>
                    <?php if ($noticeConsole !== ''): ?>
                        <pre class="console"><?= htmlspecialchars($noticeConsole) ?></pre>
                    <?php endif; ?>
                <?php endif; ?>

                <?php if ($error !== null): ?>
                    <div class="msg err"><?= htmlspecialchars($error) ?></div>
                <?php endif; ?>
            </section>

            <section class="panel">
                <h2>Historial Reciente</h2>
                <p>Se muestran solo 4 salidas: el archivo mas reciente disponible de cada marca.</p>
                <div class="history" style="margin-top: 14px;">
                    <?php if ($history === []): ?>
                        <article class="box"><strong>No hay archivos generados aun.</strong><div class="meta">La primera ejecucion aparecera aqui.</div></article>
                    <?php else: ?>
                        <?php foreach ($history as $item): ?>
                            <article class="box history-item">
                                <div>
                                    <span class="tag"><?= htmlspecialchars($historyLabel((string)$item['name'])) ?></span>
                                    <strong><?= htmlspecialchars((string)$item['name']) ?></strong>
                                    <div class="meta"><?= htmlspecialchars((string)$item['time']) ?> · <?= number_format(((int)$item['size']) / 1024, 1) ?> KB</div>
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
<script>
(() => {
    const form = document.querySelector('[data-job-form]');
    const submitButton = document.querySelector('[data-submit-button]');
    if (form && submitButton) {
        form.addEventListener('submit', () => {
            submitButton.disabled = true;
            submitButton.textContent = 'Procesando en segundo plano...';
        });
    }

    const stopForm = document.querySelector('[data-stop-form]');
    const stopButton = document.querySelector('[data-stop-button]');
    if (stopForm && stopButton && !stopButton.disabled) {
        stopForm.addEventListener('submit', (event) => {
            if (!window.confirm('Se solicitara detener todos los procesos activos de esta accion.')) {
                event.preventDefault();
                return;
            }

            stopButton.disabled = true;
            stopButton.textContent = 'Solicitando parada...';
        });
    }

    const jobBox = document.querySelector('[data-job-box]');
    if (!jobBox) {
        return;
    }

    const statusUrl = jobBox.getAttribute('data-status-url');
    const returnUrl = jobBox.getAttribute('data-return-url');
    const messageNode = jobBox.querySelector('[data-job-message]');
    const metaNode = jobBox.querySelector('[data-job-meta]');
    if (!statusUrl || !returnUrl || !messageNode || !metaNode) {
        return;
    }

    const renderMeta = (job) => {
        const status = job.status || 'running';
        const stamp = job.started_at || job.created_at || job.updated_at || '';
        metaNode.textContent = stamp ? `Estado: ${status} · ${stamp}` : `Estado: ${status}`;
    };

    const poll = async () => {
        try {
            const response = await fetch(statusUrl, { cache: 'no-store' });
            if (!response.ok) {
                throw new Error('No se pudo consultar el estado del proceso.');
            }

            const job = await response.json();
            messageNode.textContent = job.message || 'Procesando en segundo plano.';
            renderMeta(job);

            if (job.status === 'complete' || job.status === 'error' || job.status === 'cancelled') {
                window.location.href = returnUrl;
                return;
            }
        } catch (error) {
            metaNode.textContent = 'Estado: reconectando con el proceso...';
        }

        window.setTimeout(poll, 4000);
    };

    window.setTimeout(poll, 2000);
})();
</script>
</body>
</html>
