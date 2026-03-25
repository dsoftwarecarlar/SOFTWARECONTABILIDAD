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
            fn(array $job): bool => $this->isBlockingActiveJob($job)
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
                $notice = 'El proceso ya no está disponible (puede haberse limpiado). Ejecuta uno nuevo.';
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
                } elseif ($this->isExpiredActiveJob($activeJob)) {
                    $error = 'El proceso anterior excedio ' . $this->jobTimeoutLabel() . ' y se considera atascado. Inicia uno nuevo.';
                    $noticeConsole = (string)($activeJob['console'] ?? '');
                    $activeJobId = '';
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

                if ($activeJobs !== []) {
                    throw new \RuntimeException('Ya existe un proceso en ejecucion o cierre. Espera a que termine antes de iniciar otro.');
                }

                if (!is_dir($paths['template_dir'])) {
                    throw new \RuntimeException('No existe la carpeta base de plantillas mensuales.');
                }

                $brandKey = trim((string)($post['brand_key'] ?? ''));
                if ($brandKey !== '') {
                    $allowedBrands = array_map(
                        static fn(array $item): string => (string)($item['key'] ?? ''),
                        $this->config['brands'] ?? []
                    );
                    $allowedBrands = array_values(array_filter($allowedBrands, static fn(string $key): bool => $key !== ''));
                    if ($allowedBrands !== [] && !in_array($brandKey, $allowedBrands, true)) {
                        throw new \RuntimeException('La marca seleccionada no es valida.');
                    }
                }

                $uploadDefinitions = $this->uploadDefinitions();

                $storedFiles = [];
                $storeUpload = static function (array $file, string $label, array $acceptedExtensions, string $uploadsDir): string {
                    if (!is_array($file) || (int)($file['error'] ?? UPLOAD_ERR_NO_FILE) !== UPLOAD_ERR_OK) {
                        throw new \RuntimeException('Error al subir el archivo de ' . $label . '.');
                    }
                    $originalName = trim((string)($file['name'] ?? 'archivo'));
                    $tmpPath = (string)($file['tmp_name'] ?? '');
                    $extension = strtolower(pathinfo($originalName, PATHINFO_EXTENSION));
                    if (!in_array($extension, $acceptedExtensions, true)) {
                        throw new \RuntimeException('El archivo de ' . $label . ' debe ser .' . implode(' o .', $acceptedExtensions) . '.');
                    }
                    $safeBase = preg_replace('/[^A-Za-z0-9_-]+/', '_', pathinfo($originalName, PATHINFO_FILENAME));
                    $safeBase = trim((string)$safeBase, '_');
                    if ($safeBase === '') {
                        $safeBase = $label;
                    }
                    $timestamp = date('Ymd_His');
                    $inputFileName = sprintf('%s_%s.%s', $safeBase, $timestamp, $extension);
                    $inputPath = \app_join_path($uploadsDir, $inputFileName);
                    if (!move_uploaded_file($tmpPath, $inputPath)) {
                        throw new \RuntimeException('No se pudo guardar el archivo de ' . $label . '.');
                    }
                    return $inputPath;
                };

                foreach ($uploadDefinitions as $field => $meta) {
                    $required = $this->isUploadRequired($meta, $brandKey);
                    $file = $files[$field] ?? null;
                    $errorCode = is_array($file) ? (int)($file['error'] ?? UPLOAD_ERR_NO_FILE) : UPLOAD_ERR_NO_FILE;

                    if ($required && $errorCode === UPLOAD_ERR_NO_FILE) {
                        throw new \RuntimeException('Falta el archivo de ' . $meta['label'] . '.');
                    }

                    if (!is_array($file) || $errorCode === UPLOAD_ERR_NO_FILE) {
                        continue;
                    }

                    $storedFiles[$field] = $storeUpload(
                        $file,
                        (string)$meta['label'],
                        (array)($meta['accept'] ?? []),
                        $paths['uploads_dir']
                    );
                }

                $inputPath = $storedFiles['repventas_file'];

                $activeJobId = 'servicios_' . date('Ymd_His') . '_' . bin2hex(random_bytes(4));
                $createdAt = date('Y-m-d H:i:s');
                $this->jobs->write($activeJobId, [
                    'job_id' => $activeJobId,
                    'status' => 'queued',
                    'source_name' => basename($inputPath),
                    'brand_key' => $brandKey,
                    'uploads' => $storedFiles,
                    'message' => 'Proceso en cola. La pagina se actualizara automaticamente.',
                    'created_at' => $createdAt,
                ]);

                try {
                    $this->jobs->dispatch($activeJobId, $inputPath, $paths['outputs_dir'], $paths['template_dir']);
                } catch (\Throwable $dispatchException) {
                    $this->jobs->write($activeJobId, [
                        'job_id' => $activeJobId,
                        'status' => 'error',
                        'source_name' => basename($inputPath),
                        'message' => 'No se pudo iniciar el proceso en segundo plano.',
                        'error' => $dispatchException->getMessage(),
                        'uploads' => $storedFiles,
                        'brand_key' => $brandKey,
                        'created_at' => $createdAt,
                        'updated_at' => date('Y-m-d H:i:s'),
                        'completed_at' => date('Y-m-d H:i:s'),
                    ]);

                    throw new \RuntimeException('No se pudo iniciar el worker en segundo plano: ' . $dispatchException->getMessage());
                }

                $redirectUrl = \app_url('modules/cxp_servicios_marcas/index.php?job=' . rawurlencode($activeJobId));
                \header('Location: ' . $redirectUrl);
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
     * @return array<string, array<string, mixed>>
     */
    private function uploadDefinitions(): array
    {
        return [
            'px_file' => [
                'label' => 'PX',
                'accept' => ['xls', 'xlsx'],
                'scope' => 'common',
            ],
            'repventas_file' => [
                'label' => 'REP VENTAS',
                'accept' => ['xls', 'xlsx'],
                'scope' => 'common',
            ],
            'factura_changan_file' => [
                'label' => 'REP FACTURACIÓN CHANGAN',
                'accept' => ['txt'],
                'scope' => 'brand',
                'brand' => 'changan',
            ],
            'nota_changan_file' => [
                'label' => 'NOTA DE CRÉDITO CHANGAN',
                'accept' => ['txt'],
                'scope' => 'brand',
                'brand' => 'changan',
            ],
            'mayor_changan_file' => [
                'label' => 'MAYOR CHANGAN',
                'accept' => ['txt'],
                'scope' => 'brand',
                'brand' => 'changan',
            ],
            'factura_peug_file' => [
                'label' => 'REP FACTURACIÓN PEUGEOT',
                'accept' => ['txt'],
                'scope' => 'brand',
                'brand' => 'peug',
            ],
            'nota_peug_file' => [
                'label' => 'NOTA DE CRÉDITO PEUGEOT',
                'accept' => ['txt'],
                'scope' => 'brand',
                'brand' => 'peug',
            ],
            'mayor_peug_file' => [
                'label' => 'MAYOR PEUGEOT',
                'accept' => ['txt'],
                'scope' => 'brand',
                'brand' => 'peug',
            ],
            'factura_szk_file' => [
                'label' => 'REP FACTURACIÓN SUZUKI',
                'accept' => ['txt'],
                'scope' => 'brand',
                'brand' => 'szk',
            ],
            'nota_szk_file' => [
                'label' => 'NOTA DE CRÉDITO SUZUKI',
                'accept' => ['txt'],
                'scope' => 'brand',
                'brand' => 'szk',
            ],
            'mayor_szk_file' => [
                'label' => 'MAYOR SUZUKI',
                'accept' => ['txt'],
                'scope' => 'brand',
                'brand' => 'szk',
            ],
            'factura_tyt_file' => [
                'label' => 'REP FACTURACIÓN MATRIZ',
                'accept' => ['txt'],
                'scope' => 'brand',
                'brand' => 'tyt',
            ],
            'nota_tyt_file' => [
                'label' => 'NOTA DE CRÉDITO MATRIZ',
                'accept' => ['txt'],
                'scope' => 'brand',
                'brand' => 'tyt',
            ],
            'mayor_tyt_file' => [
                'label' => 'MAYOR MATRIZ',
                'accept' => ['txt'],
                'scope' => 'brand',
                'brand' => 'tyt',
            ],
        ];
    }

    /**
     * @param array<string, mixed> $definition
     */
    private function isUploadRequired(array $definition, string $brandKey): bool
    {
        $scope = (string)($definition['scope'] ?? 'common');
        if ($scope !== 'brand') {
            return true;
        }

        $definitionBrand = trim((string)($definition['brand'] ?? ''));
        if ($definitionBrand === '') {
            return false;
        }

        if ($brandKey === '') {
            return true;
        }

        return $definitionBrand === $brandKey;
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

    private function isBlockingActiveJob(array $job): bool
    {
        $status = (string)($job['status'] ?? '');
        if (!in_array($status, $this->jobs->activeStatuses(), true)) {
            return false;
        }

        return !$this->isExpiredActiveJob($job);
    }

    private function isExpiredActiveJob(array $job): bool
    {
        $status = (string)($job['status'] ?? '');
        if (!in_array($status, $this->jobs->activeStatuses(), true)) {
            return false;
        }

        $startedAt = trim((string)($job['started_at'] ?? ''));
        if ($startedAt === '') {
            return false;
        }

        $startedTs = strtotime($startedAt);
        if ($startedTs === false) {
            return false;
        }

        return (time() - $startedTs) >= $this->jobTimeoutSeconds();
    }

    private function jobTimeoutSeconds(): int
    {
        $seconds = (int)($this->config['module']['jobs']['worker_timeout_seconds'] ?? 2700);
        return $seconds > 0 ? $seconds : 2700;
    }

    private function jobTimeoutLabel(): string
    {
        $seconds = $this->jobTimeoutSeconds();
        if (($seconds % 60) === 0) {
            $minutes = intdiv($seconds, 60);
            return $minutes . ' minuto' . ($minutes === 1 ? '' : 's');
        }

        $minutes = intdiv($seconds, 60);
        $remainingSeconds = $seconds % 60;
        if ($minutes <= 0) {
            return $remainingSeconds . ' segundo' . ($remainingSeconds === 1 ? '' : 's');
        }

        return sprintf(
            '%d minuto%s %d segundo%s',
            $minutes,
            $minutes === 1 ? '' : 's',
            $remainingSeconds,
            $remainingSeconds === 1 ? '' : 's'
        );
    }
}
