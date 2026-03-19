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
    $json = json_encode($payload, JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES);
    if ($json === false) {
        throw new RuntimeException('No se pudo serializar el estado del proceso.');
    }

    $bytes = file_put_contents($path, $json, LOCK_EX);
    if ($bytes === false || $bytes < strlen($json)) {
        throw new RuntimeException('No se pudo persistir el estado del proceso.');
    }
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

function servicios_job_runtime_config(): array
{
    static $config = null;
    if (is_array($config)) {
        return $config;
    }

    $defaults = [
        'worker_timeout_seconds' => 2700,
        'stale_max_age_seconds' => 3600,
        'queued_timeout_seconds' => 300,
        'cancel_grace_seconds' => 120,
        'dispatch_boot_timeout_seconds' => 20,
    ];

    $path = app_join_path(app_root(), 'config', 'cxp', 'servicios_marcas.php');
    if (!is_file($path)) {
        $config = $defaults;
        return $config;
    }

    $loaded = require $path;
    $jobs = is_array($loaded['module']['jobs'] ?? null) ? $loaded['module']['jobs'] : [];
    $config = array_merge($defaults, $jobs);

    return $config;
}

function servicios_job_timeout_label(int $seconds): string
{
    if ($seconds <= 0) {
        return '0 minutos';
    }

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

function servicios_job_refresh_stale_jobs(string $jobsDir, ?int $maxAgeSeconds = null): void
{
    $runtime = servicios_job_runtime_config();
    if ($maxAgeSeconds === null) {
        $maxAgeSeconds = (int)($runtime['stale_max_age_seconds'] ?? 3600);
    }

    if ($maxAgeSeconds <= 0) {
        return;
    }

    $queuedTimeoutSeconds = max(60, (int)($runtime['queued_timeout_seconds'] ?? 300));
    $workerTimeoutSeconds = max(60, (int)($runtime['worker_timeout_seconds'] ?? 2700));
    $cancelGraceSeconds = max(30, (int)($runtime['cancel_grace_seconds'] ?? 120));
    $workerTimeoutLabel = servicios_job_timeout_label($workerTimeoutSeconds);
    $now = time();
    foreach (servicios_job_entries($jobsDir) as $job) {
        $status = (string)($job['status'] ?? '');
        if (!in_array($status, servicios_job_active_statuses(), true)) {
            continue;
        }

        $startedAt = trim((string)($job['started_at'] ?? ''));
        $createdAt = trim((string)($job['created_at'] ?? ''));
        $cancelRequestedAt = trim((string)($job['cancel_requested_at'] ?? ''));
        $reference = (string)($job['updated_at'] ?? $startedAt ?? $createdAt);
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

        $jobId = trim((string)($job['job_id'] ?? ''));
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

function servicios_job_resolve_php_binary(): string
{
    $candidates = [];
    $candidates[] = app_join_path(dirname(dirname(app_root())), 'php', 'php.exe');
    $candidates[] = app_join_path(dirname(dirname(app_root())), 'php', 'php');

    if (defined('PHP_BINDIR') && is_string(PHP_BINDIR) && trim(PHP_BINDIR) !== '') {
        $candidates[] = app_join_path(PHP_BINDIR, 'php.exe');
        $candidates[] = app_join_path(PHP_BINDIR, 'php');
    }

    if (defined('PHP_BINARY') && is_string(PHP_BINARY) && trim(PHP_BINARY) !== '') {
        $candidates[] = PHP_BINARY;
    }

    $seen = [];
    foreach ($candidates as $candidate) {
        $normalized = str_replace('/', '\\', (string)$candidate);
        if ($normalized === '' || isset($seen[$normalized])) {
            continue;
        }
        $seen[$normalized] = true;

        if (servicios_job_is_php_binary_candidate($candidate) && servicios_job_probe_php_binary($candidate)) {
            return $candidate;
        }
    }

    return 'php';
}

function servicios_job_is_php_binary_candidate(string $candidate): bool
{
    if (!is_file($candidate)) {
        return false;
    }

    $name = strtolower((string)pathinfo($candidate, PATHINFO_BASENAME));
    return preg_match('/^php(?:-cgi)?(?:[0-9._-]+)?(?:\.exe)?$/i', $name) === 1;
}

function servicios_job_probe_php_binary(string $binary): bool
{
    if (!is_file($binary)) {
        return false;
    }

    $output = [];
    $exitCode = 1;
    $nullDevice = DIRECTORY_SEPARATOR === '\\' ? 'NUL' : '/dev/null';
    @exec(escapeshellarg($binary) . ' -v 2>' . $nullDevice, $output, $exitCode);

    if ($exitCode !== 0) {
        return false;
    }

    if ($output === []) {
        return true;
    }

    $firstLine = strtolower(trim((string)$output[0]));
    return str_starts_with($firstLine, 'php ');
}

function servicios_job_quote_powershell_arg(string $value): string
{
    return "'" . str_replace("'", "''", $value) . "'";
}

function servicios_job_wait_until_dequeued(string $jobsDir, string $jobId, int $timeoutSeconds = 20): bool
{
    if ($timeoutSeconds <= 0) {
        return false;
    }

    $deadline = microtime(true) + $timeoutSeconds;
    while (microtime(true) < $deadline) {
        $job = servicios_job_read(servicios_job_path($jobsDir, $jobId));
        if ($job !== []) {
            $status = (string)($job['status'] ?? '');
            if ($status !== '' && $status !== 'queued') {
                return true;
            }
        }

        usleep(250000);
    }

    return false;
}

function servicios_job_terminate_process_tree(int $pid): void
{
    if ($pid <= 0) {
        return;
    }

    if (DIRECTORY_SEPARATOR === '\\') {
        exec('cmd /C taskkill /PID ' . $pid . ' /T /F >NUL 2>&1');
        return;
    }

    exec('kill -TERM ' . $pid . ' >/dev/null 2>&1');
}

/**
 * @return array<int, string>
 */
function servicios_job_read_console_lines(string $path): array
{
    if (!is_file($path)) {
        return [];
    }

    $raw = file_get_contents($path);
    if (!is_string($raw) || $raw === '') {
        return [];
    }

    // PowerShell puede escribir UTF-16 con null-bytes cuando redirige salida.
    $raw = str_replace("\0", '', $raw);
    if (strncmp($raw, "\xEF\xBB\xBF", 3) === 0) {
        $raw = substr($raw, 3);
    }
    if (strncmp($raw, "\xFF\xFE", 2) === 0 || strncmp($raw, "\xFE\xFF", 2) === 0) {
        $raw = substr($raw, 2);
    }

    $normalized = str_replace(["\r\n", "\r"], "\n", $raw);
    $lines = [];
    foreach (explode("\n", $normalized) as $line) {
        $clean = trim((string)$line);
        if ($clean !== '') {
            $lines[] = $clean;
        }
    }

    return $lines;
}

function servicios_job_dispatch(string $jobId, string $inputPath, string $outputDir, string $templateDir): void
{
    if ($jobId === '' || preg_match('/^[A-Za-z0-9_-]+$/', $jobId) !== 1) {
        throw new RuntimeException('Job invalido para iniciar en segundo plano.');
    }

    $runnerPath = app_join_path(app_root(), 'run_servicios_marcas_job.php');
    if (!is_file($runnerPath)) {
        throw new RuntimeException('No existe run_servicios_marcas_job.php para iniciar el worker.');
    }

    $phpBinary = servicios_job_resolve_php_binary();
    $workerArgs = [
        $runnerPath,
        '--job',
        $jobId,
        '--input',
        $inputPath,
        '--output-dir',
        $outputDir,
        '--template-dir',
        $templateDir,
    ];

    $jobsDir = app_ensure_dir(app_storage_path('jobs'));
    $dispatchBootTimeoutSeconds = max(5, (int)(servicios_job_runtime_config()['dispatch_boot_timeout_seconds'] ?? 20));

    if (DIRECTORY_SEPARATOR === '\\') {
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

        $psArguments = implode(', ', array_map(
            static fn(string $arg): string => servicios_job_quote_powershell_arg($arg),
            $workerArgs
        ));
        $psScript = "\$ErrorActionPreference = 'Stop'; Start-Process -FilePath "
            . servicios_job_quote_powershell_arg($phpBinary)
            . ' -ArgumentList @(' . $psArguments . ') -WindowStyle Hidden';

        $command = implode(' ', [
            escapeshellarg($powershell),
            '-Sta',
            '-NoProfile',
            '-ExecutionPolicy',
            'Bypass',
            '-Command',
            escapeshellarg($psScript),
        ]);

        $lines = [];
        $exitCode = 0;
        exec($command, $lines, $exitCode);
        if ($exitCode !== 0) {
            throw new RuntimeException('No se pudo lanzar el worker de servicios en segundo plano (exit ' . $exitCode . ').');
        }

        if (servicios_job_wait_until_dequeued($jobsDir, $jobId, $dispatchBootTimeoutSeconds)) {
            return;
        }

        $legacyCommand = 'cmd /C start "" /B ' . implode(' ', array_map(
            static fn(string $arg): string => escapeshellarg($arg),
            array_merge([$phpBinary], $workerArgs)
        )) . ' >NUL 2>&1';

        $legacyHandle = @popen($legacyCommand, 'r');
        if ($legacyHandle !== false) {
            pclose($legacyHandle);
        } else {
            $legacyExitCode = 0;
            exec($legacyCommand, $lines, $legacyExitCode);
            if ($legacyExitCode !== 0) {
                throw new RuntimeException(
                    'No se pudo iniciar el worker en segundo plano (PowerShell y fallback cmd fallaron).'
                );
            }
        }

        if (!servicios_job_wait_until_dequeued($jobsDir, $jobId, $dispatchBootTimeoutSeconds)) {
            throw new RuntimeException(
                'El worker no inicio y el proceso quedo en cola. Verifica permisos de PHP/PowerShell e intenta de nuevo. ' .
                'PHP detectado: ' . $phpBinary . '. Job: ' . $jobId . '.'
            );
        }

        return;
    }

    $workerCommand = implode(' ', array_map(static fn(string $part): string => escapeshellarg($part), $workerArgs));
    $lines = [];
    $exitCode = 0;
    exec(escapeshellarg($phpBinary) . ' ' . $workerCommand . ' > /dev/null 2>&1 &', $lines, $exitCode);
    if ($exitCode !== 0) {
        throw new RuntimeException('No se pudo lanzar el worker de servicios en segundo plano (exit ' . $exitCode . ').');
    }

    if (!servicios_job_wait_until_dequeued($jobsDir, $jobId, $dispatchBootTimeoutSeconds)) {
        throw new RuntimeException(
            'El worker no inicio y el proceso quedo en cola. Intenta nuevamente. ' .
            'PHP detectado: ' . $phpBinary . '. Job: ' . $jobId . '.'
        );
    }
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

    $console = '';

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

        $brandKey = trim((string)($existingJob['brand_key'] ?? ''));
        $uploads = is_array($existingJob['uploads'] ?? null) ? $existingJob['uploads'] : [];
        $getUpload = static function (string $key) use ($uploads): string {
            $path = (string)($uploads[$key] ?? '');
            return is_file($path) ? $path : '';
        };
        $command = implode(' ', [
            escapeshellarg($powershell),
            '-Sta',
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
            '-BrandKey',
            escapeshellarg($brandKey),
            '-PxPath',
            escapeshellarg($getUpload('px_file')),
            '-RepVtasPath',
            escapeshellarg($getUpload('repventas_file')),
            '-FacturaChanganPath',
            escapeshellarg($getUpload('factura_changan_file')),
            '-NotaChanganPath',
            escapeshellarg($getUpload('nota_changan_file')),
            '-MayorChanganPath',
            escapeshellarg($getUpload('mayor_changan_file')),
            '-FacturaPeugPath',
            escapeshellarg($getUpload('factura_peug_file')),
            '-NotaPeugPath',
            escapeshellarg($getUpload('nota_peug_file')),
            '-MayorPeugPath',
            escapeshellarg($getUpload('mayor_peug_file')),
            '-FacturaSzkPath',
            escapeshellarg($getUpload('factura_szk_file')),
            '-NotaSzkPath',
            escapeshellarg($getUpload('nota_szk_file')),
            '-MayorSzkPath',
            escapeshellarg($getUpload('mayor_szk_file')),
            '-FacturaTytPath',
            escapeshellarg($getUpload('factura_tyt_file')),
            '-NotaTytPath',
            escapeshellarg($getUpload('nota_tyt_file')),
            '-MayorTytPath',
            escapeshellarg($getUpload('mayor_tyt_file')),
        ]);

        $outputLines = [];
        $configByKey = servicios_job_output_config();
        $startedTs = time();
        $lastHeartbeatAt = $startedTs;
        $timedOut = false;
        $workerTimeoutSeconds = max(60, (int)(servicios_job_runtime_config()['worker_timeout_seconds'] ?? 2700));
        $workerTimeoutLabel = servicios_job_timeout_label($workerTimeoutSeconds);
        $consoleLogPath = app_join_path($jobsDir, 'servicios_marcas_' . $jobId . '.log');
        if (is_file($consoleLogPath)) {
            @unlink($consoleLogPath);
        }

        $consoleStream = fopen($consoleLogPath, 'ab');
        if (!is_resource($consoleStream)) {
            throw new RuntimeException('No se pudo abrir el log temporal del worker de servicios por marca.');
        }

        $touchRunningStatus = static function (string $message) use ($jobPath, $writeStatus): void {
            $currentStatus = (string)(servicios_job_read($jobPath)['status'] ?? '');
            if ($currentStatus === 'cancel_requested') {
                return;
            }

            $writeStatus('running', ['message' => $message]);
        };

        $process = proc_open(
            $command,
            [
                1 => $consoleStream,
                2 => $consoleStream,
            ],
            $pipes
        );

        if (!is_resource($process)) {
            fclose($consoleStream);
            throw new RuntimeException('No se pudo lanzar el worker de servicios por marca.');
        }

        while (true) {
            $statusInfo = proc_get_status($process);
            $isRunning = (bool)($statusInfo['running'] ?? false);
            if (!$isRunning) {
                break;
            }

            if ((time() - $startedTs) >= $workerTimeoutSeconds) {
                servicios_job_terminate_process_tree((int)($statusInfo['pid'] ?? 0));
                $timedOut = true;
                break;
            }

            if ((time() - $lastHeartbeatAt) >= 20) {
                $elapsed = max(1, time() - $startedTs);
                $minutes = intdiv($elapsed, 60);
                $seconds = $elapsed % 60;
                $touchRunningStatus(sprintf(
                    'Procesando plantillas en segundo plano. Tiempo transcurrido: %02d:%02d.',
                    $minutes,
                    $seconds
                ));
                $lastHeartbeatAt = time();
            }

            usleep(500000);
        }

        $exitCode = proc_close($process);
        fclose($consoleStream);

        if ($timedOut) {
            throw new RuntimeException(
                'El worker de servicios supero el tiempo maximo permitido (' . $workerTimeoutLabel . ') y fue detenido para evitar bloqueos.'
            );
        }

        $outputLines = servicios_job_read_console_lines($consoleLogPath);
        if (is_file($consoleLogPath)) {
            @unlink($consoleLogPath);
        }

        $console = trim(implode(PHP_EOL, $outputLines));
        $downloads = [];
        $summary = [];
        $cancelMessage = '';

        foreach ($outputLines as $line) {
            $text = trim((string)$line);
            $cancelPos = strpos($text, 'CANCELLED|');
            if ($cancelPos !== false) {
                $text = substr($text, $cancelPos);
                $cancelMessage = trim(substr($text, strlen('CANCELLED|')));
                continue;
            }

            $infoPos = strpos($text, 'INFO|');
            if ($infoPos !== false) {
                $text = substr($text, $infoPos);
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

            $outputPos = strpos($text, 'OUTPUT|');
            if ($outputPos === false) {
                continue;
            }
            $text = substr($text, $outputPos);

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

        // Fallback: si la salida de consola no se pudo parsear, resolver por artefactos reales del run_stamp.
        if ($downloads === []) {
            foreach (glob(app_join_path($outputDir, 'servicios_*_' . $runStamp . '.xls')) ?: [] as $generatedPath) {
                if (!is_file($generatedPath)) {
                    continue;
                }

                $generatedName = basename($generatedPath);
                $generatedKey = strtolower(pathinfo($generatedName, PATHINFO_FILENAME));
                if (isset($downloads[$generatedKey])) {
                    continue;
                }

                $label = 'SALIDA';
                foreach ($configByKey as $brandKey => $brandConfig) {
                    if (str_starts_with(strtolower($generatedName), 'servicios_' . strtolower($brandKey) . '_')) {
                        $label = (string)($brandConfig['label'] ?? strtoupper($brandKey));
                        break;
                    }
                }

                $downloads[$generatedKey] = [
                    'label' => $label,
                    'name' => $generatedName,
                    'download_url' => app_output_download_url($generatedName),
                ];
            }
        }

        if ($downloads === []) {
            foreach (glob(app_join_path($outputDir, 'servicios_*.xls')) ?: [] as $generatedPath) {
                if (!is_file($generatedPath)) {
                    continue;
                }

                $mtime = filemtime($generatedPath);
                if ($mtime === false || $mtime < ($startedTs - 10)) {
                    continue;
                }

                $generatedName = basename($generatedPath);
                $generatedKey = strtolower(pathinfo($generatedName, PATHINFO_FILENAME));
                if (isset($downloads[$generatedKey])) {
                    continue;
                }

                $label = 'SALIDA';
                foreach ($configByKey as $brandKey => $brandConfig) {
                    if (str_starts_with(strtolower($generatedName), 'servicios_' . strtolower($brandKey) . '_')) {
                        $label = (string)($brandConfig['label'] ?? strtoupper($brandKey));
                        break;
                    }
                }

                $downloads[$generatedKey] = [
                    'label' => $label,
                    'name' => $generatedName,
                    'download_url' => app_output_download_url($generatedName),
                ];
            }
        }

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
            'console' => $console,
            'completed_at' => date('Y-m-d H:i:s'),
        ]);
        throw $exception;
    }
}
