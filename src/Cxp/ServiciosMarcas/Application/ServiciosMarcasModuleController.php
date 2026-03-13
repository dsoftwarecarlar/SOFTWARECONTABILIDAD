<?php
declare(strict_types=1);

namespace App\Cxp\ServiciosMarcas\Application;

use App\Cxp\ServiciosMarcas\Domain\HistoryLabelResolver;
use App\Cxp\ServiciosMarcas\Infrastructure\ServiciosMarcasJobGateway;

final class ServiciosMarcasModuleController
{
    private ServiciosMarcasJobGateway $jobs;
    private HistoryLabelResolver $historyLabels;
    /**
     * @var array<string, mixed>
     */
    private array $config;

    /**
     * @param array<string, mixed> $config
     */
    public function __construct(
        ServiciosMarcasJobGateway $jobs,
        HistoryLabelResolver $historyLabels,
        array $config
    ) {
        $this->jobs = $jobs;
        $this->historyLabels = $historyLabels;
        $this->config = $config;
    }

    /**
     * @param array<string, mixed> $server
     * @param array<string, mixed> $get
     * @param array<string, mixed> $post
     * @param array<string, mixed> $files
     * @return array<string, mixed>
     */
    public function handle(array $server, array $get, array $post, array $files): array
    {
        $context = \app_workspace_module('cxp', 'cxp_servicios_marcas');
        if ($context === null) {
            throw new \RuntimeException('No se encontro la configuracion del modulo.');
        }

        $windowContext = \app_workspace_window('cxp', 'conciliacion_servicios_marcas');
        $workspace = $context['workspace'];
        $currentModule = $context['module'];
        $window = $windowContext['window'] ?? null;
        $paths = $this->resolvePaths();
        $requestMethod = (string)($server['REQUEST_METHOD'] ?? 'GET');

        $result = null;
        $error = null;
        $notice = null;
        $noticeConsole = '';
        $activeJobId = '';
        $activeJob = null;
        $pendingJob = null;

        $this->jobs->refreshStaleJobs();
        $jobEntries = $this->jobs->entries();
        $activeJobs = array_values(array_filter(
            $jobEntries,
            fn(array $job): bool => in_array((string)($job['status'] ?? ''), $this->jobs->activeStatuses(), true)
        ));

        $stoppedCount = isset($get['stopped'])
            ? filter_var($get['stopped'], FILTER_VALIDATE_INT, FILTER_NULL_ON_FAILURE)
            : null;
        if ($stoppedCount !== null) {
            $notice = $stoppedCount > 0
                ? 'Se solicito detener ' . $stoppedCount . ' proceso(s) activo(s).'
                : 'No habia procesos activos para detener.';
        }

        $statusJobId = trim((string)($get['status'] ?? ''));
        if ($statusJobId !== '') {
            $job = $this->jobs->read($statusJobId);
            \header('Content-Type: application/json; charset=utf-8');

            if ($job === null) {
                \http_response_code(404);
                echo json_encode(['status' => 'missing', 'message' => 'No se encontro el trabajo solicitado.'], JSON_UNESCAPED_SLASHES);
                exit;
            }

            echo json_encode($job, JSON_UNESCAPED_SLASHES);
            exit;
        }

        $activeJobId = trim((string)($get['job'] ?? ''));
        if ($activeJobId !== '') {
            $activeJob = $this->jobs->read($activeJobId);
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
                $action = trim((string)($post['action'] ?? 'process'));
                if ($action === 'stop_all') {
                    $stopResult = $this->jobs->requestCancelAll();
                    $redirectParams = ['stopped' => (string)$stopResult['count']];
                    $returnJobId = trim((string)($post['return_job'] ?? ''));
                    if ($returnJobId !== '') {
                        $redirectParams['job'] = $returnJobId;
                    }

                    \header('Location: ' . \app_url('modules/cxp_servicios_marcas/index.php?' . http_build_query($redirectParams)));
                    exit;
                }

                \ignore_user_abort(true);
                \set_time_limit(0);

                if (!is_dir($paths['template_dir'])) {
                    throw new \RuntimeException('No existe la carpeta base de plantillas mensuales.');
                }

                if (!isset($files['excel_file'])) {
                    throw new \RuntimeException('No se recibio el archivo Excel.');
                }

                $file = $files['excel_file'];
                if (!is_array($file) || (int)($file['error'] ?? UPLOAD_ERR_NO_FILE) !== UPLOAD_ERR_OK) {
                    throw new \RuntimeException('Error al subir el archivo Excel.');
                }

                $originalName = trim((string)($file['name'] ?? 'reporte.xls'));
                $tmpPath = (string)($file['tmp_name'] ?? '');
                $extension = strtolower(pathinfo($originalName, PATHINFO_EXTENSION));
                $acceptedExtensions = $this->acceptedExtensions();
                if (!in_array($extension, $acceptedExtensions, true)) {
                    throw new \RuntimeException('Solo se permiten archivos Excel .' . implode(' o .', $acceptedExtensions) . '.');
                }

                $safeBase = preg_replace('/[^A-Za-z0-9_-]+/', '_', pathinfo($originalName, PATHINFO_FILENAME));
                $safeBase = trim((string)$safeBase, '_');
                if ($safeBase === '') {
                    $safeBase = 'reporte_servicios';
                }

                $timestamp = date('Ymd_His');
                $inputFileName = sprintf('%s_%s.%s', $safeBase, $timestamp, $extension);
                $inputPath = \app_join_path($paths['uploads_dir'], $inputFileName);

                if (!move_uploaded_file($tmpPath, $inputPath)) {
                    throw new \RuntimeException('No se pudo guardar el Excel subido.');
                }

                $activeJobId = 'servicios_' . date('Ymd_His') . '_' . bin2hex(random_bytes(4));
                $this->jobs->write($activeJobId, [
                    'job_id' => $activeJobId,
                    'status' => 'queued',
                    'source_name' => $originalName,
                    'message' => 'Proceso en cola. La pagina se actualizara automaticamente.',
                    'created_at' => date('Y-m-d H:i:s'),
                ]);

                $redirectUrl = \app_url('modules/cxp_servicios_marcas/index.php?job=' . rawurlencode($activeJobId));
                \header('Location: ' . $redirectUrl);
                \header('Content-Length: 0');
                \header('Connection: close');
                if (function_exists('session_write_close')) {
                    \session_write_close();
                }
                while (ob_get_level() > 0) {
                    @ob_end_flush();
                }
                \flush();

                try {
                    $this->jobs->run($activeJobId, $inputPath, $paths['outputs_dir'], $paths['template_dir']);
                } catch (\Throwable) {
                    // El estado del job queda persistido para el polling del cliente.
                }
                exit;
            } catch (\Throwable $exception) {
                $error = $exception->getMessage();
            }
        }

        return [
            'brand' => \app_brand(),
            'workspace' => $workspace,
            'currentModule' => $currentModule,
            'window' => $window,
            'templateDir' => $paths['template_dir'],
            'result' => $result,
            'error' => $error,
            'notice' => $notice,
            'noticeConsole' => $noticeConsole,
            'activeJobId' => $activeJobId,
            'activeJob' => $activeJob,
            'pendingJob' => $pendingJob,
            'activeJobs' => $activeJobs,
            'history' => $this->buildHistory(),
            'historyLabelResolver' => $this->historyLabels,
            'pageConfig' => $this->config['module'] ?? [],
            'brands' => $this->config['brands'] ?? [],
            'stylesheets' => [
                \app_asset_url('css/base.css'),
                \app_asset_url('css/pages/cxp-servicios-marcas.css'),
            ],
            'scripts' => [
                \app_asset_url('js/cxp-servicios-marcas.js'),
            ],
        ];
    }

    /**
     * @return array<string, string>
     */
    private function resolvePaths(): array
    {
        return [
            'uploads_dir' => \app_ensure_dir(\app_storage_path('uploads')),
            'outputs_dir' => \app_ensure_dir(\app_storage_path('outputs')),
            'jobs_dir' => $this->jobs->jobsDir(),
            'template_dir' => (string)($this->config['paths']['template_dir'] ?? ''),
        ];
    }

    /**
     * @return string[]
     */
    private function acceptedExtensions(): array
    {
        $extensions = $this->config['module']['accepted_extensions'] ?? ['xls', 'xlsx'];
        return array_values(array_map(static fn($extension): string => strtolower((string)$extension), $extensions));
    }

    /**
     * @return array<int, array<string, mixed>>
     */
    private function buildHistory(): array
    {
        $history = [];
        $seenHistoryBrands = [];
        $scanLimit = (int)($this->config['module']['history_scan_limit'] ?? 20);
        $historyLimit = (int)($this->config['module']['history_limit'] ?? 4);

        foreach (\app_list_output_files_for_action('servicios', $scanLimit) as $item) {
            $brandKey = $this->historyLabels->detectBrandKey((string)($item['name'] ?? ''));
            if ($brandKey === null || isset($seenHistoryBrands[$brandKey])) {
                continue;
            }

            $seenHistoryBrands[$brandKey] = true;
            $history[] = $item;
            if (count($history) >= $historyLimit) {
                break;
            }
        }

        return $history;
    }
}
