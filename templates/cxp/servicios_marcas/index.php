<?php
declare(strict_types=1);

/** @var array<string, mixed> $brand */
/** @var array<string, mixed> $workspace */
/** @var array<string, mixed> $currentModule */
/** @var array<string, mixed>|null $window */
/** @var string $templateDir */
/** @var array<string, mixed>|null $result */
/** @var string|null $error */
/** @var string|null $notice */
/** @var string $noticeConsole */
/** @var string $activeJobId */
/** @var array<string, mixed>|null $pendingJob */
/** @var array<int, array<string, mixed>> $activeJobs */
/** @var array<int, array<string, mixed>> $history */
/** @var App\Cxp\ServiciosMarcas\Domain\HistoryLabelResolver $historyLabelResolver */
/** @var array<string, mixed> $pageConfig */
/** @var array<int, array<string, string>> $brands */
/** @var array<int, string> $stylesheets */
/** @var array<int, string> $scripts */
?>
<!doctype html>
<html lang="es">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title><?= htmlspecialchars((string)$workspace['title']) ?> | <?= htmlspecialchars((string)$currentModule['title']) ?></title>
    <?php foreach ($stylesheets as $stylesheet): ?>
        <link rel="stylesheet" href="<?= htmlspecialchars($stylesheet) ?>">
    <?php endforeach; ?>
</head>
<body class="app-page servicios-marcas-page">
<div class="page">
    <div class="topbar">
        <div>
            <span class="chip"><?= htmlspecialchars((string)$brand['platform']) ?></span>
            <strong><?= htmlspecialchars((string)$workspace['title']) ?></strong>
            <span class="note"><?= htmlspecialchars((string)($pageConfig['module_note'] ?? '')) ?></span>
        </div>
        <a class="back-link" href="<?= htmlspecialchars((string)($window['url'] ?? ($workspace['home_url'] ?? app_url()))) ?>">Volver a la ventana</a>
    </div>

    <section class="hero">
        <div class="hero-grid">
            <div>
                <span class="chip"><?= htmlspecialchars((string)($pageConfig['hero']['chip'] ?? 'Proceso')) ?></span>
                <h1><?= htmlspecialchars((string)$currentModule['title']) ?></h1>
                <p class="lead"><?= htmlspecialchars((string)($pageConfig['hero']['lead'] ?? '')) ?></p>
                <div class="hero-note"><?= htmlspecialchars((string)($pageConfig['hero']['note'] ?? '')) ?></div>
            </div>
            <div class="hero-side">
                <?php if (!empty($brand['logo_exists'])): ?>
                    <div class="logo-stage">
                        <div class="logo-frame">
                            <img src="<?= htmlspecialchars((string)$brand['logo_url']) ?>" alt="<?= htmlspecialchars((string)$brand['company']) ?>">
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
            <h2><?= htmlspecialchars((string)($pageConfig['template_panel']['title'] ?? 'Plantillas')) ?></h2>
            <p><?= htmlspecialchars((string)($pageConfig['template_panel']['description'] ?? '')) ?></p>
            <div class="brands">
                <article class="box">
                    <span class="tag"><?= htmlspecialchars((string)($pageConfig['template_panel']['base_tag'] ?? 'Base')) ?></span>
                    <strong><?= htmlspecialchars((string)($pageConfig['template_panel']['base_title'] ?? 'Plantillas')) ?></strong>
                    <p><?= htmlspecialchars((string)($pageConfig['template_panel']['base_description'] ?? '')) ?></p>
                    <div class="role-note template-note">Estas plantillas solo sirven como base visual y de validacion. No son la salida que se descarga al terminar.</div>
                    <div class="base-path"><?= htmlspecialchars($templateDir) ?></div>
                </article>
                <article class="box">
                    <?php foreach ($brands as $brandItem): ?>
                        <div class="brand-row">
                            <span class="tag"><?= htmlspecialchars((string)$brandItem['label']) ?></span>
                            <span><?= htmlspecialchars((string)$brandItem['description']) ?></span>
                        </div>
                    <?php endforeach; ?>
                </article>
            </div>
        </aside>

        <main class="stack">
            <section class="panel">
                <h2><?= htmlspecialchars((string)($pageConfig['upload_panel']['title'] ?? 'Cargar Excel')) ?></h2>
                <p><?= htmlspecialchars((string)($pageConfig['upload_panel']['description'] ?? '')) ?></p>
                <div class="form-grid">
                    <form method="post" enctype="multipart/form-data" data-job-form>
                        <input type="hidden" name="action" value="process">
                        <input type="file" name="excel_file" accept=".xls,.xlsx" required>
                        <button type="submit" data-submit-button data-processing-label="<?= htmlspecialchars((string)($pageConfig['upload_panel']['processing_label'] ?? 'Procesando...')) ?>">
                            <?= htmlspecialchars((string)($pageConfig['upload_panel']['button_label'] ?? 'Procesar')) ?>
                        </button>
                    </form>
                    <article class="box">
                        <span class="tag"><?= htmlspecialchars((string)($pageConfig['upload_panel']['retention_tag'] ?? 'Retencion')) ?></span>
                        <strong><?= htmlspecialchars((string)($pageConfig['upload_panel']['retention_title'] ?? '')) ?></strong>
                        <p><?= htmlspecialchars((string)($pageConfig['upload_panel']['retention_description'] ?? '')) ?></p>
                        <div class="role-note output-note">Los archivos finales aparecen en esta pantalla y en `storage/outputs`. No salen desde la carpeta de plantillas base.</div>
                    </article>
                </div>
                <div class="control-row" style="margin-top: 12px;">
                    <form method="post" data-stop-form>
                        <input type="hidden" name="action" value="stop_all">
                        <?php if ($activeJobId !== ''): ?>
                            <input type="hidden" name="return_job" value="<?= htmlspecialchars($activeJobId) ?>">
                        <?php endif; ?>
                        <button
                            type="submit"
                            class="stop-button"
                            data-stop-button
                            data-confirm-message="<?= htmlspecialchars((string)($pageConfig['stop']['confirm_message'] ?? '')) ?>"
                            data-processing-label="<?= htmlspecialchars((string)($pageConfig['stop']['processing_label'] ?? 'Procesando...')) ?>"
                            <?= $activeJobs === [] ? 'disabled' : '' ?>
                        >
                            <?= htmlspecialchars((string)($pageConfig['stop']['button_label'] ?? 'Detener')) ?>
                        </button>
                    </form>
                    <div class="meta"><?= htmlspecialchars(sprintf((string)($pageConfig['stop']['meta_template'] ?? 'Procesos activos: %d'), count($activeJobs))) ?></div>
                </div>

                <?php if ($pendingJob !== null): ?>
                    <div
                        class="msg pending"
                        data-job-box
                        data-job-id="<?= htmlspecialchars($activeJobId) ?>"
                        data-status-url="<?= htmlspecialchars(app_url('modules/cxp_servicios_marcas/index.php?status=' . rawurlencode($activeJobId))) ?>"
                        data-return-url="<?= htmlspecialchars(app_url('modules/cxp_servicios_marcas/index.php?job=' . rawurlencode($activeJobId))) ?>"
                        data-poll-interval="<?= htmlspecialchars((string)($pageConfig['poll']['interval_ms'] ?? 4000)) ?>"
                        data-poll-delay="<?= htmlspecialchars((string)($pageConfig['poll']['initial_delay_ms'] ?? 2000)) ?>"
                        data-reconnecting-message="<?= htmlspecialchars((string)($pageConfig['poll']['reconnecting_message'] ?? '')) ?>"
                        data-status-error="<?= htmlspecialchars((string)($pageConfig['poll']['status_error'] ?? '')) ?>"
                    >
                        <strong data-job-message><?= htmlspecialchars((string)($pendingJob['message'] ?? 'Procesando en segundo plano.')) ?></strong>
                        <div class="meta" data-job-meta>
                            Estado: <?= htmlspecialchars((string)($pendingJob['status'] ?? 'running')) ?>
                            <?php if (!empty($pendingJob['started_at'])): ?>
                                - Inicio: <?= htmlspecialchars((string)$pendingJob['started_at']) ?>
                            <?php elseif (!empty($pendingJob['created_at'])): ?>
                                - Creado: <?= htmlspecialchars((string)$pendingJob['created_at']) ?>
                            <?php endif; ?>
                        </div>
                        <div class="meta"><?= htmlspecialchars((string)($pageConfig['upload_panel']['pending_hint'] ?? '')) ?></div>
                    </div>
                <?php endif; ?>

                <?php if ($result !== null): ?>
                    <div class="msg ok">Archivo procesado: <?= htmlspecialchars((string)$result['source_name']) ?>. Los enlaces de abajo son salidas generadas del upload actual.</div>
                    <div class="result-guide">
                        <article class="box output-box">
                            <span class="tag">Entrada subida</span>
                            <strong><?= htmlspecialchars((string)$result['source_name']) ?></strong>
                            <div class="meta">Fuente del proceso actual.</div>
                        </article>
                        <article class="box output-box">
                            <span class="tag">Salida generada</span>
                            <strong><?= count((array)($result['downloads'] ?? [])) ?> archivo(s)</strong>
                            <div class="meta">Descarga estos enlaces, no la plantilla base del lateral.</div>
                        </article>
                    </div>
                    <?php if (($result['summary'] ?? []) !== []): ?>
                        <div class="summary-grid" style="margin-top: 12px;">
                            <?php foreach ((array)$result['summary'] as $item): ?>
                                <article class="box">
                                    <span class="tag"><?= htmlspecialchars((string)($item['label'] ?? 'MARCA')) ?></span>
                                    <strong><?= (int)($item['rows'] ?? 0) ?> filas procesadas</strong>
                                    <div class="meta">
                                        Facturas fallback: <?= (int)($item['invoice_fallbacks'] ?? 0) ?>
                                        - Notas fallback: <?= (int)($item['note_fallbacks'] ?? 0) ?>
                                    </div>
                                </article>
                            <?php endforeach; ?>
                        </div>
                    <?php endif; ?>
                    <div class="downloads" style="margin-top: 12px;">
                        <?php foreach ((array)$result['downloads'] as $download): ?>
                            <article class="box download-card">
                                <div>
                                    <span class="tag">Salida generada · <?= htmlspecialchars((string)$download['label']) ?></span>
                                    <strong><?= htmlspecialchars((string)$download['name']) ?></strong>
                                    <div class="meta">Archivo creado desde el Excel subido en este proceso.</div>
                                </div>
                                <a class="button-link" href="<?= htmlspecialchars((string)$download['download_url']) ?>">Descargar</a>
                            </article>
                        <?php endforeach; ?>
                    </div>
                    <?php if (($result['console'] ?? '') !== ''): ?>
                        <pre class="console"><?= htmlspecialchars((string)$result['console']) ?></pre>
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
                <h2><?= htmlspecialchars((string)($pageConfig['history_panel']['title'] ?? 'Historial')) ?></h2>
                <p><?= htmlspecialchars((string)($pageConfig['history_panel']['description'] ?? '')) ?> Aqui solo se listan salidas generadas; las plantillas base viven en la columna lateral.</p>
                <div class="history" style="margin-top: 14px;">
                    <?php if ($history === []): ?>
                        <article class="box">
                            <strong><?= htmlspecialchars((string)($pageConfig['history_panel']['empty_title'] ?? 'Sin archivos')) ?></strong>
                            <div class="meta"><?= htmlspecialchars((string)($pageConfig['history_panel']['empty_description'] ?? '')) ?></div>
                        </article>
                    <?php else: ?>
                        <?php foreach ($history as $item): ?>
                            <article class="box history-item">
                                <div>
                                    <span class="tag"><?= htmlspecialchars($historyLabelResolver->resolveLabel((string)$item['name'])) ?></span>
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
<?php foreach ($scripts as $script): ?>
    <script src="<?= htmlspecialchars($script) ?>"></script>
<?php endforeach; ?>
</body>
</html>
