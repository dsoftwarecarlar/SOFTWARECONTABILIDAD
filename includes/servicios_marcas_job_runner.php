<?php
declare(strict_types=1);

require_once __DIR__ . '/bootstrap.php';

function servicios_job_path(string $jobsDir, string $jobId): string
{
    return app_join_path($jobsDir, 'servicios_marcas_' . $jobId . '.json');
}

function servicios_job_cancel_path(string $jobsDir, string $jobId): string
{
    return app_join_path($jobsDir, 'servicios_marcas_' . $jobId . '.stop');
}

function servicios_job_read(string $path): array
{
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

function servicios_job_write(string $path, array $payload): void
{
    file_put_contents(
        $path,
        json_encode($payload, JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES),
        LOCK_EX
    );
}

function servicios_job_cleanup(string $jobsDir, int $keep = 24): void
{
    if ($keep < 0 || !is_dir($jobsDir)) {
        return;
    }

    $files = glob(app_join_path($jobsDir, 'servicios_marcas_*.json')) ?: [];
    usort(
        $files,
        static function (string $left, string $right): int {
            $timeLeft = filectime($left) ?: (filemtime($left) ?: 0);
            $timeRight = filectime($right) ?: (filemtime($right) ?: 0);
            if ($timeLeft === $timeRight) {
                return strcmp(basename($right), basename($left));
            }
            return $timeRight <=> $timeLeft;
        }
    );

    foreach (array_slice($files, $keep) as $file) {
        if (is_file($file)) {
            $jobId = preg_replace('/^servicios_marcas_|\.json$/', '', basename($file));
            $cancelPath = servicios_job_cancel_path($jobsDir, (string)$jobId);
            @unlink($file);
            if (is_file($cancelPath)) {
                @unlink($cancelPath);
            }
        }
    }

    foreach (glob(app_join_path($jobsDir, 'servicios_marcas_*.stop')) ?: [] as $cancelFile) {
        $jobId = preg_replace('/^servicios_marcas_|\.stop$/', '', basename($cancelFile));
        $jobPath = servicios_job_path($jobsDir, (string)$jobId);
        if (!is_file($jobPath)) {
            @unlink($cancelFile);
        }
    }
}

function servicios_job_output_config(): array
{
    return [
        'changan' => ['label' => 'CHANGAN', 'prefix' => 'servicios_changan_'],
        'peug' => ['label' => 'PEUGEOT', 'prefix' => 'servicios_peug_'],
        'szk' => ['label' => 'SUZUKI', 'prefix' => 'servicios_szk_'],
        'tyt' => ['label' => 'MATRIZ', 'prefix' => 'servicios_tyt_'],
    ];
}

function servicios_job_entries(string $jobsDir): array
{
    $files = glob(app_join_path($jobsDir, 'servicios_marcas_*.json')) ?: [];
    usort(
        $files,
        static function (string $left, string $right): int {
            $timeLeft = filectime($left) ?: (filemtime($left) ?: 0);
            $timeRight = filectime($right) ?: (filemtime($right) ?: 0);
            if ($timeLeft === $timeRight) {
                return strcmp(basename($right), basename($left));
            }
            return $timeRight <=> $timeLeft;
        }
    );

    $entries = [];
    foreach ($files as $file) {
        $data = servicios_job_read($file);
        if ($data === []) {
            continue;
        }

        $entries[] = array_merge(
            $data,
            [
                '_path' => $file,
                '_cancel_path' => servicios_job_cancel_path(
                    $jobsDir,
                    preg_replace('/^servicios_marcas_|\.json$/', '', basename($file)) ?: ''
                ),
            ]
        );
    }

    return $entries;
}

function servicios_job_active_statuses(): array
{
    return ['queued', 'running', 'cancel_requested'];
}

function servicios_job_refresh_stale_jobs(string $jobsDir, int $maxAgeSeconds = 3600): void
{
    if ($maxAgeSeconds <= 0) {
        return;
    }

    $now = time();
    foreach (servicios_job_entries($jobsDir) as $job) {
        $status = (string)($job['status'] ?? '');
        if (!in_array($status, servicios_job_active_statuses(), true)) {
            continue;
        }

        $startedAt = trim((string)($job['started_at'] ?? ''));
        $createdAt = trim((string)($job['created_at'] ?? ''));
        $reference = (string)($job['updated_at'] ?? $startedAt ?? $createdAt);
        $referenceTs = strtotime($reference);
        $createdTs = strtotime($createdAt);
        $shouldCloseUnstarted = $startedAt === ''
            && $createdTs !== false
            && ($now - $createdTs) >= 300;

        if (!$shouldCloseUnstarted && ($referenceTs === false || ($now - $referenceTs) < $maxAgeSeconds)) {
            continue;
        }

        $jobId = trim((string)($job['job_id'] ?? ''));
        if ($jobId === '') {
            continue;
        }

        $payload = $job;
        unset($payload['_path'], $payload['_cancel_path']);
        $payload['status'] = 'cancelled';
        $payload['message'] = 'Proceso antiguo cerrado automaticamente para depurar la cola.';
        $payload['updated_at'] = date('Y-m-d H:i:s');
        $payload['completed_at'] = date('Y-m-d H:i:s');
        servicios_job_write(servicios_job_path($jobsDir, $jobId), $payload);

        $cancelPath = servicios_job_cancel_path($jobsDir, $jobId);
        if (is_file($cancelPath)) {
            @unlink($cancelPath);
        }
    }
}

function servicios_job_request_cancel_all(string $jobsDir): array
{
    servicios_job_refresh_stale_jobs($jobsDir);

    $count = 0;
    $jobIds = [];
    $timestamp = date('Y-m-d H:i:s');

    foreach (servicios_job_entries($jobsDir) as $job) {
        $status = (string)($job['status'] ?? '');
        if (!in_array($status, servicios_job_active_statuses(), true)) {
            continue;
        }

        $jobId = trim((string)($job['job_id'] ?? ''));
        if ($jobId === '') {
            continue;
        }

        $cancelPath = (string)($job['_cancel_path'] ?? servicios_job_cancel_path($jobsDir, $jobId));
        if (!is_file($cancelPath)) {
            file_put_contents($cancelPath, $timestamp . PHP_EOL, LOCK_EX);
        }

        $payload = $job;
        unset($payload['_path'], $payload['_cancel_path']);
        $payload['status'] = 'cancel_requested';
        $payload['message'] = 'Se solicito detener el proceso. Esperando cierre del script.';
        $payload['cancel_requested_at'] = $timestamp;
        $payload['updated_at'] = $timestamp;
        servicios_job_write(servicios_job_path($jobsDir, $jobId), $payload);

        $count++;
        $jobIds[] = $jobId;
    }

    return [
        'count' => $count,
        'job_ids' => $jobIds,
    ];
}

function servicios_job_run(string $jobId, string $inputPath, string $outputDir, string $templateDir): void
{
    if ($jobId === '' || preg_match('/^[A-Za-z0-9_-]+$/', $jobId) !== 1) {
        throw new RuntimeException('Job invalido.');
    }

    $jobsDir = app_ensure_dir(app_storage_path('jobs'));
    $jobPath = servicios_job_path($jobsDir, $jobId);
    $cancelPath = servicios_job_cancel_path($jobsDir, $jobId);
    $existingJob = servicios_job_read($jobPath);
    $sourceName = trim((string)($existingJob['source_name'] ?? ''));
    $runStamp = trim((string)($existingJob['run_stamp'] ?? ''));
    if ($sourceName === '') {
        $sourceName = basename($inputPath);
    }
    if ($runStamp === '') {
        $runStamp = date('Ymd_His') . '_' . substr(bin2hex(random_bytes(2)), 0, 4);
    }

    $writeStatus = static function (string $status, array $extra = []) use ($jobPath, $jobId, $sourceName, $runStamp): void {
        $current = servicios_job_read($jobPath);
        $payload = array_merge(
            $current,
            [
                'job_id' => $jobId,
                'status' => $status,
                'source_name' => $sourceName,
                'run_stamp' => $runStamp,
                'updated_at' => date('Y-m-d H:i:s'),
            ],
            $extra
        );

        servicios_job_write($jobPath, $payload);
    };

    $deleteGeneratedOutputs = static function () use ($outputDir, $runStamp): void {
        foreach (glob(app_join_path($outputDir, 'servicios_*_' . $runStamp . '.xls')) ?: [] as $path) {
            if (is_file($path)) {
                @unlink($path);
            }
        }
    };

    try {
        if (!is_file($inputPath)) {
            throw new RuntimeException('No se encontro el archivo Excel subido.');
        }

        if (!is_dir($outputDir)) {
            app_ensure_dir($outputDir);
        }

        if (!is_dir($outputDir)) {
            throw new RuntimeException('No se encontro la carpeta de salidas.');
        }

        if (!is_dir($templateDir)) {
            throw new RuntimeException('No se encontro la carpeta de plantillas.');
        }

        $scriptPath = app_join_path(app_root(), 'scripts', 'cxp', 'servicios_marcas', 'run.ps1');
        if (!is_file($scriptPath)) {
            throw new RuntimeException('No existe el worker de servicios por marca en scripts/cxp/servicios_marcas/run.ps1.');
        }

        $existingStatus = (string)($existingJob['status'] ?? '');
        if (is_file($cancelPath) || in_array($existingStatus, ['cancel_requested', 'cancelled'], true)) {
            if (is_file($cancelPath)) {
                @unlink($cancelPath);
            }

            $writeStatus('cancelled', [
                'message' => 'Proceso detenido por solicitud del usuario.',
                'downloads' => [],
                'summary' => [],
                'console' => '',
                'completed_at' => date('Y-m-d H:i:s'),
            ]);
            return;
        }

        $writeStatus('running', [
            'message' => 'Procesando plantillas en segundo plano. Este paso puede tardar varios minutos.',
            'started_at' => date('Y-m-d H:i:s'),
        ]);

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
            '-InputPath',
            escapeshellarg($inputPath),
            '-OutputDir',
            escapeshellarg($outputDir),
            '-TemplateDir',
            escapeshellarg($templateDir),
            '-RunStamp',
            escapeshellarg($runStamp),
            '-CancelPath',
            escapeshellarg($cancelPath),
            '2>&1',
        ]);

        $outputLines = [];
        $exitCode = 0;
        exec($command, $outputLines, $exitCode);

        $console = trim(implode(PHP_EOL, $outputLines));
        $downloads = [];
        $summary = [];
        $cancelMessage = '';

        foreach ($outputLines as $line) {
            $text = trim((string)$line);
            if (str_starts_with($text, 'CANCELLED|')) {
                $cancelMessage = trim(substr($text, strlen('CANCELLED|')));
                continue;
            }

            if (preg_match('/^INFO\|processing\|([a-z0-9_]+)\|rows=(\d+)$/i', $text, $matches) === 1) {
                $key = strtolower($matches[1]);
                $summary[$key] = array_merge(
                    $summary[$key] ?? [],
                    ['key' => $key, 'rows' => (int)$matches[2]]
                );
                continue;
            }

            if (preg_match('/^INFO\|([a-z0-9_]+)\|invoice_fallbacks=(\d+)\|note_fallbacks=(\d+)$/i', $text, $matches) === 1) {
                $key = strtolower($matches[1]);
                $summary[$key] = array_merge(
                    $summary[$key] ?? [],
                    [
                        'key' => $key,
                        'invoice_fallbacks' => (int)$matches[2],
                        'note_fallbacks' => (int)$matches[3],
                    ]
                );
                continue;
            }

            if (!str_starts_with($text, 'OUTPUT|')) {
                continue;
            }

            $parts = explode('|', $text, 3);
            if (count($parts) < 3) {
                continue;
            }

            [, $fileName, $label] = $parts;
            $fileName = trim($fileName);
            $key = strtolower(pathinfo($fileName, PATHINFO_FILENAME));
            if (isset($downloads[$key])) {
                continue;
            }

            $downloads[$key] = [
                'label' => trim($label),
                'name' => $fileName,
                'download_url' => app_output_download_url($fileName),
            ];
        }

        $configByKey = servicios_job_output_config();
        $order = array_keys($configByKey);
        $downloads = array_values($downloads);
        usort(
            $downloads,
            static function (array $left, array $right) use ($order): int {
                $leftName = strtolower((string)($left['name'] ?? ''));
                $rightName = strtolower((string)($right['name'] ?? ''));

                $leftIndex = count($order);
                foreach ($order as $position => $key) {
                    if (str_starts_with($leftName, 'servicios_' . $key . '_')) {
                        $leftIndex = $position;
                        break;
                    }
                }

                $rightIndex = count($order);
                foreach ($order as $position => $key) {
                    if (str_starts_with($rightName, 'servicios_' . $key . '_')) {
                        $rightIndex = $position;
                        break;
                    }
                }

                if ($leftIndex === $rightIndex) {
                    return strcmp($leftName, $rightName);
                }

                return $leftIndex <=> $rightIndex;
            }
        );

        $summaryList = [];
        foreach ($order as $key) {
            if (!isset($summary[$key])) {
                continue;
            }

            $summaryList[] = [
                'key' => $key,
                'label' => (string)($configByKey[$key]['label'] ?? strtoupper($key)),
                'rows' => (int)($summary[$key]['rows'] ?? 0),
                'invoice_fallbacks' => (int)($summary[$key]['invoice_fallbacks'] ?? 0),
                'note_fallbacks' => (int)($summary[$key]['note_fallbacks'] ?? 0),
            ];
        }

        if ($cancelMessage !== '') {
            $deleteGeneratedOutputs();
            if (is_file($cancelPath)) {
                @unlink($cancelPath);
            }

            $writeStatus('cancelled', [
                'message' => $cancelMessage,
                'downloads' => [],
                'summary' => $summaryList,
                'console' => $console,
                'completed_at' => date('Y-m-d H:i:s'),
            ]);
            return;
        }

        if ($exitCode !== 0) {
            $deleteGeneratedOutputs();
            throw new RuntimeException($console !== '' ? $console : 'El script de servicios termino con error.');
        }

        if ($downloads === []) {
            throw new RuntimeException('El proceso termino sin generar archivos de salida.');
        }

        foreach ($configByKey as $config) {
            $prefix = strtolower((string)$config['prefix']);
            app_cleanup_output_files(
                app_output_retention_limit(),
                static fn(string $name): bool => str_starts_with(strtolower($name), $prefix)
            );
        }

        app_cleanup_upload_files(app_upload_retention_limit());
        servicios_job_cleanup($jobsDir);
        if (is_file($cancelPath)) {
            @unlink($cancelPath);
        }

        $writeStatus('complete', [
            'message' => 'Proceso terminado. Ya puedes descargar las plantillas generadas.',
            'downloads' => $downloads,
            'summary' => $summaryList,
            'console' => $console,
            'completed_at' => date('Y-m-d H:i:s'),
        ]);
    } catch (Throwable $exception) {
        servicios_job_cleanup($jobsDir);
        if (is_file($cancelPath)) {
            @unlink($cancelPath);
        }
        $writeStatus('error', [
            'message' => 'El proceso termino con error.',
            'error' => $exception->getMessage(),
            'completed_at' => date('Y-m-d H:i:s'),
        ]);
        throw $exception;
    }
}
