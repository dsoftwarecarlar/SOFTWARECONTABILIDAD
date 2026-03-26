<?php

declare(strict_types=1);

namespace App\Services;

use App\Support\AppUrl;
use Illuminate\Http\UploadedFile;
use RuntimeException;

final class ServiciosMarcasRunner
{
    private const ACTIVE_STATUSES = ['queued', 'running', 'cancel_requested'];
    private const MAYOR_VENTAS_PATTERNS = [
        'changan' => '/04\.01\.01\.12\.(0001|0002|0003|0010|0012|0014)/',
        'peug' => '/04\.01\.01\.13\.(0001|0002|0003|0010|0012|0014)/',
        'szk' => '/04\.01\.01\.14\.(0001|0002|0003|0010|0012|0014)/',
        'tyt' => '/04\.01\.01\.11\.(0001|0002|0003|0010|0012|0014)/',
    ];
    private const FLEXIBLE_MAYOR_VENTAS_BRANDS = ['tyt'];
    private const GENERIC_MAYOR_VENTAS_PATTERN = '/04\.01\.01\.\d{2}\.\d{4}/';

    public function __construct(
        private ExternalProcessService $processes,
        private PythonBridge $python,
        private ServiciosMarcasDispatcher $dispatcher
    ) {
    }

    public function isManagedModule(string $moduleSlug): bool
    {
        return $moduleSlug === 'servicios-marcas';
    }

    /**
     * @return array<string, mixed>
     */
    public function moduleConfig(): array
    {
        $config = config('cxp.servicios_module');
        if (!is_array($config)) {
            throw new RuntimeException('No existe configuracion Laravel para Servicios por Marca.');
        }

        return $config;
    }

    public function refreshStaleJobs(): void
    {
        $maxAgeSeconds = max(60, (int) (($this->moduleConfig()['jobs']['stale_max_age_seconds'] ?? 3600)));
        $queuedTimeoutSeconds = max(60, (int) (($this->moduleConfig()['jobs']['queued_timeout_seconds'] ?? 300)));
        $workerTimeoutSeconds = max(60, (int) (($this->moduleConfig()['jobs']['worker_timeout_seconds'] ?? 2700)));
        $cancelGraceSeconds = max(30, (int) (($this->moduleConfig()['jobs']['cancel_grace_seconds'] ?? 120)));
        $workerTimeoutLabel = $this->jobTimeoutLabel();
        $now = time();

        foreach ($this->jobEntries() as $job) {
            $status = (string) ($job['status'] ?? '');
            if (!in_array($status, self::ACTIVE_STATUSES, true)) {
                continue;
            }

            $startedAt = trim((string) ($job['started_at'] ?? ''));
            $createdAt = trim((string) ($job['created_at'] ?? ''));
            $cancelRequestedAt = trim((string) ($job['cancel_requested_at'] ?? ''));
            $reference = (string) ($job['updated_at'] ?? $startedAt ?? $createdAt);
            $referenceTs = strtotime($reference);
            $startedTs = strtotime($startedAt);
            $createdTs = strtotime($createdAt);
            $cancelRequestedTs = strtotime($cancelRequestedAt);

            $shouldCloseUnstarted = $startedAt === ''
                && $createdTs !== false
                && ($now - $createdTs) >= $queuedTimeoutSeconds;
            $shouldCloseRunningTimeout = $status === 'running'
                && $startedTs !== false
                && ($now - $startedTs) >= $workerTimeoutSeconds;
            $shouldCloseCancelRequested = $status === 'cancel_requested'
                && (
                    ($cancelRequestedTs !== false && ($now - $cancelRequestedTs) >= $cancelGraceSeconds)
                    || ($referenceTs !== false && ($now - $referenceTs) >= $cancelGraceSeconds)
                );

            if (
                !$shouldCloseUnstarted
                && !$shouldCloseRunningTimeout
                && !$shouldCloseCancelRequested
                && ($referenceTs === false || ($now - $referenceTs) < $maxAgeSeconds)
            ) {
                continue;
            }

            $jobId = trim((string) ($job['job_id'] ?? ''));
            if ($jobId === '') {
                continue;
            }

            $payload = $job;
            unset($payload['_path'], $payload['_cancel_path']);
            $payload['status'] = 'cancelled';
            if ($shouldCloseCancelRequested) {
                $payload['message'] = 'Proceso cancelado automaticamente tras confirmar cierre del worker.';
            } elseif ($shouldCloseRunningTimeout) {
                $payload['message'] = 'Proceso cerrado automaticamente por exceder ' . $workerTimeoutLabel . ' sin finalizar.';
            } else {
                $payload['message'] = 'Proceso antiguo cerrado automaticamente para depurar la cola.';
            }
            $payload['updated_at'] = date('Y-m-d H:i:s');
            $payload['completed_at'] = date('Y-m-d H:i:s');
            $this->writeJobSnapshot($jobId, $payload);

            $cancelPath = (string) ($job['_cancel_path'] ?? $this->cancelPath($jobId));
            if (is_file($cancelPath)) {
                @unlink($cancelPath);
            }
        }
    }

    /**
     * @return array<int, array<string, mixed>>
     */
    public function activeJobs(): array
    {
        $this->refreshStaleJobs();

        return array_values(array_filter(
            $this->jobEntries(),
            fn (array $job): bool => $this->isBlockingActiveJob($job)
        ));
    }

    /**
     * @return array<string, mixed>|null
     */
    public function readJob(string $jobId): ?array
    {
        if ($jobId === '' || preg_match('/^[A-Za-z0-9_-]+$/', $jobId) !== 1) {
            return null;
        }

        $data = $this->readJobSnapshot($jobId);
        if ($data === []) {
            return null;
        }

        return $this->normalizeJob($data);
    }

    /**
     * @param array<string, UploadedFile|array<int, UploadedFile>|null> $uploadedFiles
     */
    public function dispatch(array $uploadedFiles, string $brandKey = ''): string
    {
        if ($this->activeJobs() !== []) {
            throw new RuntimeException('Ya existe un proceso en ejecucion o cierre. Espera a que termine antes de iniciar otro.');
        }

        $config = $this->moduleConfig();
        $templateDir = (string) ($config['template_dir'] ?? '');
        if (!is_dir($templateDir)) {
            throw new RuntimeException('No existe la carpeta base de plantillas mensuales.');
        }

        $allowedBrands = array_values(array_filter(array_map(
            static fn (array $item): string => (string) ($item['key'] ?? ''),
            (array) ($config['brands'] ?? [])
        )));
        if ($brandKey !== '' && $allowedBrands !== [] && !in_array($brandKey, $allowedBrands, true)) {
            throw new RuntimeException('La marca seleccionada no es valida.');
        }

        $storedFiles = $this->persistUploadedInputs($uploadedFiles, $brandKey, $config);
        $this->validateMayorVentasUploads($storedFiles, $config, $brandKey);
        $inputPath = (string) ($storedFiles['repventas_file'] ?? '');
        if ($inputPath === '' || !is_file($inputPath)) {
            throw new RuntimeException('No se pudo preparar el archivo REP VENTAS para el worker.');
        }

        $jobId = 'servicios_' . date('Ymd_His') . '_' . bin2hex(random_bytes(4));
        $createdAt = date('Y-m-d H:i:s');

        $this->writeJobSnapshot($jobId, [
            'job_id' => $jobId,
            'status' => 'queued',
            'source_name' => basename($inputPath),
            'brand_key' => $brandKey,
            'uploads' => $storedFiles,
            'message' => 'Proceso en cola. La pagina se actualizara automaticamente.',
            'created_at' => $createdAt,
        ]);

        try {
            $workerArguments = [
                'worker_path' => base_path('../python_services/processors/servicios_marcas/worker.py'),
                'input_path' => $inputPath,
                'output_dir' => (string) config('cxp.storage.outputs'),
                'template_dir' => $templateDir,
                'jobs_dir' => $this->jobsDir(),
                'worker_timeout_seconds' => (int) (($config['jobs']['worker_timeout_seconds'] ?? 2700)),
                'cancel_grace_seconds' => (int) (($config['jobs']['cancel_grace_seconds'] ?? 120)),
                'queued_timeout_seconds' => (int) (($config['jobs']['queued_timeout_seconds'] ?? 300)),
                'dispatch_boot_timeout_seconds' => (int) (($config['jobs']['dispatch_boot_timeout_seconds'] ?? 20)),
            ];

            $this->python->execute(
                (string) ($config['python_processor'] ?? 'servicios_marcas.dispatch'),
                [],
                $this->bridgeResultPath($jobId),
                null,
                [
                    'job_id' => $jobId,
                    ...$workerArguments,
                ]
            );
            $this->dispatcher->dispatch(
                $jobId,
                $workerArguments,
                $this->jobsDir(),
                max(5, (int) (($config['jobs']['dispatch_boot_timeout_seconds'] ?? 20)))
            );
        } catch (\Throwable $exception) {
            $this->writeJobSnapshot($jobId, [
                'job_id' => $jobId,
                'status' => 'error',
                'source_name' => basename($inputPath),
                'message' => 'No se pudo iniciar el proceso en segundo plano.',
                'error' => $exception->getMessage(),
                'uploads' => $storedFiles,
                'brand_key' => $brandKey,
                'created_at' => $createdAt,
                'updated_at' => date('Y-m-d H:i:s'),
                'completed_at' => date('Y-m-d H:i:s'),
            ]);

            throw new RuntimeException('No se pudo iniciar el worker en segundo plano: ' . $exception->getMessage(), 0, $exception);
        }

        return $jobId;
    }

    /**
     * @return array{count:int, job_ids: array<int, string>}
     */
    public function requestCancelAll(): array
    {
        $this->refreshStaleJobs();

        $count = 0;
        $jobIds = [];
        $timestamp = date('Y-m-d H:i:s');

        foreach ($this->jobEntries() as $job) {
            $status = (string) ($job['status'] ?? '');
            if (!in_array($status, self::ACTIVE_STATUSES, true)) {
                continue;
            }

            $jobId = trim((string) ($job['job_id'] ?? ''));
            if ($jobId === '') {
                continue;
            }

            $cancelPath = (string) ($job['_cancel_path'] ?? $this->cancelPath($jobId));
            if (!is_file($cancelPath)) {
                file_put_contents($cancelPath, $timestamp . PHP_EOL, LOCK_EX);
            }

            $payload = $job;
            unset($payload['_path'], $payload['_cancel_path']);
            $payload['status'] = 'cancel_requested';
            $payload['message'] = 'Se solicito detener el proceso. Esperando cierre del script.';
            $payload['cancel_requested_at'] = $timestamp;
            $payload['updated_at'] = $timestamp;
            $this->writeJobSnapshot($jobId, $payload);

            $count++;
            $jobIds[] = $jobId;
        }

        return [
            'count' => $count,
            'job_ids' => $jobIds,
        ];
    }

    /**
     * @return array<int, array<string, mixed>>
     */
    public function history(): array
    {
        $config = $this->moduleConfig();
        $scanLimit = max(1, (int) ($config['history_scan_limit'] ?? 20));
        $historyLimit = max(1, (int) ($config['history_limit'] ?? 4));
        $outputConfig = (array) ($config['output_config'] ?? []);

        $files = [];
        foreach (glob((string) config('cxp.storage.outputs') . DIRECTORY_SEPARATOR . 'servicios_*.xls') ?: [] as $file) {
            if (!is_file($file)) {
                continue;
            }

            $timestamp = filectime($file) ?: (filemtime($file) ?: time());
            $files[] = [
                'path' => $file,
                'name' => basename($file),
                'size' => (int) (filesize($file) ?: 0),
                'timestamp' => $timestamp,
                'time' => date('Y-m-d H:i:s', $timestamp),
            ];
        }

        usort(
            $files,
            static function (array $left, array $right): int {
                if ($left['timestamp'] === $right['timestamp']) {
                    return strcmp((string) $right['name'], (string) $left['name']);
                }

                return ((int) $right['timestamp']) <=> ((int) $left['timestamp']);
            }
        );

        $history = [];
        $seenBrands = [];
        foreach (array_slice($files, 0, $scanLimit) as $item) {
            $brandKey = $this->detectBrandKey((string) $item['name'], $outputConfig);
            if ($brandKey === null || isset($seenBrands[$brandKey])) {
                continue;
            }

            $seenBrands[$brandKey] = true;
            $item['label'] = (string) ($outputConfig[$brandKey]['label'] ?? strtoupper($brandKey));
            $item['download_url'] = AppUrl::route('downloads.show', ['file' => $item['name']]);
            $history[] = $item;
            if (count($history) >= $historyLimit) {
                break;
            }
        }

        return $history;
    }

    public function jobTimeoutLabel(): string
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

    /**
     * @return array{runtime: string, command: string}
     */
    public function commandPreview(): array
    {
        return [
            'runtime' => 'python',
            'command' => sprintf(
                '%s %s <manifest:%s> -> %s',
                (string) config('python.binary', 'python'),
                base_path('../python_services/cli.py'),
                (string) ($this->moduleConfig()['python_processor'] ?? 'servicios_marcas.dispatch'),
                base_path('../python_services/processors/servicios_marcas/worker.py')
            ),
        ];
    }

    /**
     * @return array{runtime: string, command: string}
     */
    public function workerCommand(): array
    {
        return $this->processes->buildCommand(
            'powershell+excel-com',
            base_path('../scripts/cxp/servicios_marcas/run.ps1')
        );
    }

    /**
     * @return array<string, mixed>|null
     */
    public function pendingOrResult(string $jobId): ?array
    {
        return $this->readJob($jobId);
    }

    /**
     * @return array{runtime: string, command: string}
     */
    public function legacyCommand(): array
    {
        return $this->commandPreview();
    }

    private function jobsDir(): string
    {
        return (string) config('cxp.storage.jobs');
    }

    private function bridgeResultPath(string $jobId): string
    {
        $directory = storage_path('app/python_results');
        if (!is_dir($directory) && !mkdir($directory, 0775, true) && !is_dir($directory)) {
            throw new RuntimeException('No se pudo preparar storage/app/python_results para Servicios por Marca.');
        }

        return $directory . DIRECTORY_SEPARATOR . 'servicios_dispatch_' . $jobId . '.json';
    }


    private function jobTimeoutSeconds(): int
    {
        $config = $this->moduleConfig();
        $seconds = (int) (($config['jobs']['worker_timeout_seconds'] ?? 2700));

        return $seconds > 0 ? $seconds : 2700;
    }

    /**
     * @return array<int, array<string, mixed>>
     */
    private function jobEntries(): array
    {
        $files = glob($this->jobsDir() . DIRECTORY_SEPARATOR . 'servicios_marcas_*.json') ?: [];
        usort(
            $files,
            static function (string $left, string $right): int {
                $timeLeft = self::safeFileTimestamp($left);
                $timeRight = self::safeFileTimestamp($right);
                if ($timeLeft === $timeRight) {
                    return strcmp(basename($right), basename($left));
                }

                return $timeRight <=> $timeLeft;
            }
        );

        $entries = [];
        foreach ($files as $file) {
            if (!is_file($file)) {
                continue;
            }

            $jobId = preg_replace('/^servicios_marcas_|\.json$/', '', basename($file)) ?: '';
            if ($jobId === '') {
                continue;
            }

            $data = $this->readJobSnapshot((string) $jobId);
            if ($data === []) {
                continue;
            }

            $entries[] = array_merge($data, [
                '_path' => $file,
                '_cancel_path' => $this->cancelPath((string) $jobId),
            ]);
        }

        return $entries;
    }

    private function isBlockingActiveJob(array $job): bool
    {
        $status = (string) ($job['status'] ?? '');
        if (!in_array($status, $this->activeStatuses(), true)) {
            return false;
        }

        return !$this->isExpiredActiveJob($job);
    }

    private function isExpiredActiveJob(array $job): bool
    {
        $status = (string) ($job['status'] ?? '');
        if (!in_array($status, $this->activeStatuses(), true)) {
            return false;
        }

        $startedAt = trim((string) ($job['started_at'] ?? ''));
        if ($startedAt === '') {
            return false;
        }

        $startedTs = strtotime($startedAt);
        if ($startedTs === false) {
            return false;
        }

        return (time() - $startedTs) >= $this->jobTimeoutSeconds();
    }

    /**
     * @return array<int, string>
     */
    private function activeStatuses(): array
    {
        return self::ACTIVE_STATUSES;
    }

    /**
     * @param array<string, UploadedFile|array<int, UploadedFile>|null> $uploadedFiles
     * @param array<string, mixed> $config
     * @return array<string, string>
     */
    private function persistUploadedInputs(array $uploadedFiles, string $brandKey, array $config): array
    {
        $uploadsDir = (string) config('cxp.storage.uploads');
        if (!is_dir($uploadsDir) && !mkdir($uploadsDir, 0775, true) && !is_dir($uploadsDir)) {
            throw new RuntimeException('No se pudo preparar storage/uploads para Servicios por Marca.');
        }

        $storedFiles = [];
        foreach ((array) ($config['upload_definitions'] ?? []) as $field => $meta) {
            if (!is_array($meta)) {
                continue;
            }

            $required = $this->isUploadRequired($meta, $brandKey);
            $uploaded = $uploadedFiles[$field] ?? null;
            if (!$uploaded instanceof UploadedFile) {
                if ($required) {
                    throw new RuntimeException('Falta el archivo de ' . (string) ($meta['label'] ?? $field) . '.');
                }
                continue;
            }

            $originalName = trim((string) $uploaded->getClientOriginalName());
            $extension = strtolower((string) $uploaded->getClientOriginalExtension());
            $accepted = array_map(static fn ($item): string => strtolower((string) $item), (array) ($meta['accept'] ?? []));
            if ($accepted !== [] && !in_array($extension, $accepted, true)) {
                throw new RuntimeException(
                    'El archivo de ' . (string) ($meta['label'] ?? $field) . ' debe ser .' . implode(' o .', $accepted) . '.'
                );
            }

            $safeBase = preg_replace('/[^A-Za-z0-9_-]+/', '_', pathinfo($originalName, PATHINFO_FILENAME));
            $safeBase = trim((string) $safeBase, '_');
            if ($safeBase === '') {
                $safeBase = (string) $field;
            }

            $timestamp = date('Ymd_His');
            $inputFileName = sprintf('%s_%s.%s', $safeBase, $timestamp, $extension);
            $uploaded->move($uploadsDir, $inputFileName);
            $storedFiles[$field] = $uploadsDir . DIRECTORY_SEPARATOR . $inputFileName;
        }

        return $storedFiles;
    }

    /**
     * @param array<string, string> $storedFiles
     * @param array<string, mixed> $config
     */
    private function validateMayorVentasUploads(array $storedFiles, array $config, string $brandKey): void
    {
        foreach ((array) ($config['upload_definitions'] ?? []) as $field => $meta) {
            if (!is_array($meta)) {
                continue;
            }

            $scope = (string) ($meta['scope'] ?? 'common');
            if ($scope !== 'brand') {
                continue;
            }

            $definitionBrand = trim((string) ($meta['brand'] ?? ''));
            if ($definitionBrand === '' || !str_contains((string) $field, 'mayor_')) {
                continue;
            }

            if ($brandKey !== '' && $definitionBrand !== $brandKey) {
                continue;
            }

            $path = (string) ($storedFiles[$field] ?? '');
            if ($path === '' || !is_file($path)) {
                continue;
            }

            $this->assertMayorVentasUploadMatchesBrand(
                $definitionBrand,
                $path,
                (string) ($meta['label'] ?? $field)
            );
        }
    }

    private function assertMayorVentasUploadMatchesBrand(string $brandKey, string $path, string $label): void
    {
        $pattern = $this->expectedMayorVentasPattern($brandKey);
        if ($pattern === null) {
            return;
        }

        $contents = @file_get_contents($path);
        if ($contents === false || trim($contents) === '') {
            throw new RuntimeException('El archivo de ' . $label . ' esta vacio o no se pudo leer.');
        }

        if (preg_match($pattern, $contents) === 1) {
            return;
        }

        $firstDetectedAccount = '';
        if (preg_match('/\d{2}\.\d{2}\.\d{2}\.\d{2}\.\d{4}/', $contents, $matches) === 1) {
            $firstDetectedAccount = (string) ($matches[0] ?? '');
        }

        $suffix = $firstDetectedAccount !== ''
            ? ' Se detecto la cuenta ' . $firstDetectedAccount . '.'
            : '';

        throw new RuntimeException(
            'El archivo de ' . $label . ' no corresponde al MAYOR VENTAS de la marca seleccionada.' .
            $suffix .
            ' Debe incluir cuentas 04.01.01.xx.xxxx.'
        );
    }

    private function expectedMayorVentasPattern(string $brandKey): ?string
    {
        if (in_array($brandKey, self::FLEXIBLE_MAYOR_VENTAS_BRANDS, true)) {
            return self::GENERIC_MAYOR_VENTAS_PATTERN;
        }

        return self::MAYOR_VENTAS_PATTERNS[$brandKey] ?? null;
    }

    /**
     * @param array<string, mixed> $meta
     */
    private function isUploadRequired(array $meta, string $brandKey): bool
    {
        $scope = (string) ($meta['scope'] ?? 'common');
        if ($scope !== 'brand') {
            return true;
        }

        $definitionBrand = trim((string) ($meta['brand'] ?? ''));
        if ($definitionBrand === '') {
            return false;
        }

        if ($brandKey === '') {
            return true;
        }

        return $definitionBrand === $brandKey;
    }

    /**
     * @param array<string, mixed> $job
     * @return array<string, mixed>
     */
    private function normalizeJob(array $job): array
    {
        $downloads = [];
        foreach ((array) ($job['downloads'] ?? []) as $download) {
            if (!is_array($download)) {
                continue;
            }

            $name = trim((string) ($download['name'] ?? ''));
            if ($name === '') {
                continue;
            }

            $downloads[] = [
                'label' => (string) ($download['label'] ?? 'SALIDA'),
                'name' => $name,
                'download_url' => AppUrl::route('downloads.show', ['file' => $name]),
            ];
        }

        $job['downloads'] = $downloads;

        return $job;
    }

    /**
     * @param array<string, array{label:string, prefix:string}> $outputConfig
     */
    private function detectBrandKey(string $fileName, array $outputConfig): ?string
    {
        $name = strtolower($fileName);
        foreach ($outputConfig as $key => $config) {
            $prefix = strtolower((string) ($config['prefix'] ?? ''));
            if ($prefix !== '' && str_starts_with($name, $prefix)) {
                return (string) $key;
            }
        }

        return null;
    }

    private function jobPath(string $jobId): string
    {
        return $this->jobsDir() . DIRECTORY_SEPARATOR . 'servicios_marcas_' . $jobId . '.json';
    }

    private function cancelPath(string $jobId): string
    {
        return $this->jobsDir() . DIRECTORY_SEPARATOR . 'servicios_marcas_' . $jobId . '.stop';
    }

    /**
     * @return array<string, mixed>
     */
    private function readJobSnapshot(string $jobId): array
    {
        $path = $this->jobPath($jobId);
        if (!is_file($path)) {
            return [];
        }

        $raw = file_get_contents($path);
        if ($raw === false || trim($raw) === '') {
            return [];
        }

        $data = json_decode($raw, true);
        return is_array($data) ? $data : [];
    }

    /**
     * @param array<string, mixed> $payload
     */
    private function writeJobSnapshot(string $jobId, array $payload): void
    {
        $json = json_encode($payload, JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES);
        if ($json === false) {
            throw new RuntimeException('No se pudo serializar el estado del proceso.');
        }

        $bytes = file_put_contents($this->jobPath($jobId), $json, LOCK_EX);
        if ($bytes === false || $bytes < strlen($json)) {
            throw new RuntimeException('No se pudo persistir el estado del proceso.');
        }
    }

    private static function safeFileTimestamp(string $path): int
    {
        if (!is_file($path)) {
            return 0;
        }

        $ctime = @filectime($path);
        if (is_int($ctime) && $ctime > 0) {
            return $ctime;
        }

        $mtime = @filemtime($path);
        if (is_int($mtime) && $mtime > 0) {
            return $mtime;
        }

        return 0;
    }
}
