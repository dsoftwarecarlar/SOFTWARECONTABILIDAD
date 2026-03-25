<?php

declare(strict_types=1);

namespace App\Services;

use RuntimeException;

final class ServiciosMarcasDispatcher
{
    public function dispatch(
        string $jobId,
        array $workerArguments,
        string $jobsDir,
        int $bootTimeoutSeconds
    ): void {
        if ($jobId === '' || preg_match('/^[A-Za-z0-9_-]+$/', $jobId) !== 1) {
            throw new RuntimeException('Job invalido para iniciar en segundo plano.');
        }

        $workerPath = (string) ($workerArguments['worker_path'] ?? '');
        if (!is_file($workerPath)) {
            throw new RuntimeException('No existe el worker Python de servicios para iniciar el proceso.');
        }

        $pythonBinary = $this->resolvePythonBinary();
        $workerArgs = [
            $workerPath,
            '--job',
            $jobId,
            '--input',
            (string) ($workerArguments['input_path'] ?? ''),
            '--output-dir',
            (string) ($workerArguments['output_dir'] ?? ''),
            '--template-dir',
            (string) ($workerArguments['template_dir'] ?? ''),
            '--jobs-dir',
            $jobsDir,
            '--worker-timeout-seconds',
            (string) ($workerArguments['worker_timeout_seconds'] ?? 2700),
            '--cancel-grace-seconds',
            (string) ($workerArguments['cancel_grace_seconds'] ?? 120),
            '--queued-timeout-seconds',
            (string) ($workerArguments['queued_timeout_seconds'] ?? 300),
            '--dispatch-boot-timeout-seconds',
            (string) ($workerArguments['dispatch_boot_timeout_seconds'] ?? 20),
        ];

        if (DIRECTORY_SEPARATOR === '\\') {
            $this->dispatchOnWindows($pythonBinary, $workerArgs, $jobsDir, $jobId, $bootTimeoutSeconds);
            return;
        }

        $workerCommand = implode(' ', array_map(static fn (string $part): string => escapeshellarg($part), $workerArgs));
        $lines = [];
        $exitCode = 0;
        exec(escapeshellarg($pythonBinary) . ' ' . $workerCommand . ' > /dev/null 2>&1 &', $lines, $exitCode);
        if ($exitCode !== 0) {
            throw new RuntimeException('No se pudo lanzar el worker de servicios en segundo plano (exit ' . $exitCode . ').');
        }

        if (!$this->waitUntilDequeued($jobsDir, $jobId, $bootTimeoutSeconds)) {
            throw new RuntimeException(
                'El worker no inicio y el proceso quedo en cola. Intenta nuevamente. ' .
                'Python detectado: ' . $pythonBinary . '. Job: ' . $jobId . '.'
            );
        }
    }

    /**
     * @param list<string> $workerArgs
     */
    private function dispatchOnWindows(
        string $pythonBinary,
        array $workerArgs,
        string $jobsDir,
        string $jobId,
        int $bootTimeoutSeconds
    ): void {
        $powershell = implode(DIRECTORY_SEPARATOR, [
            getenv('WINDIR') ?: 'C:\\Windows',
            'System32',
            'WindowsPowerShell',
            'v1.0',
            'powershell.exe',
        ]);
        if (!is_file($powershell)) {
            $powershell = 'powershell.exe';
        }

        $psArguments = implode(', ', array_map(
            fn (string $arg): string => $this->quotePowerShellArg($arg),
            $workerArgs
        ));
        $psScript = "\$ErrorActionPreference = 'Stop'; Start-Process -FilePath "
            . $this->quotePowerShellArg($pythonBinary)
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

        if ($this->waitUntilDequeued($jobsDir, $jobId, $bootTimeoutSeconds)) {
            return;
        }

        $legacyCommand = 'cmd /C start "" /B ' . implode(' ', array_map(
            static fn (string $arg): string => escapeshellarg($arg),
            array_merge([$pythonBinary], $workerArgs)
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

        if (!$this->waitUntilDequeued($jobsDir, $jobId, $bootTimeoutSeconds)) {
            throw new RuntimeException(
                'El worker no inicio y el proceso quedo en cola. Verifica permisos de Python/PowerShell e intenta de nuevo. ' .
                'Python detectado: ' . $pythonBinary . '. Job: ' . $jobId . '.'
            );
        }
    }

    private function resolvePythonBinary(): string
    {
        $configured = trim((string) config('python.binary', 'python'));
        if ($configured === '') {
            return 'python';
        }

        if ((str_contains($configured, '\\') || str_contains($configured, '/')) && !is_file($configured)) {
            throw new RuntimeException('No existe el binario Python configurado para Servicios por Marca.');
        }

        return $configured;
    }

    private function quotePowerShellArg(string $value): string
    {
        return "'" . str_replace("'", "''", $value) . "'";
    }

    private function waitUntilDequeued(string $jobsDir, string $jobId, int $timeoutSeconds): bool
    {
        if ($timeoutSeconds <= 0) {
            return false;
        }

        $deadline = microtime(true) + $timeoutSeconds;
        while (microtime(true) < $deadline) {
            $job = $this->readJobSnapshot($jobsDir, $jobId);
            if ($job !== []) {
                $status = (string) ($job['status'] ?? '');
                if ($status !== '' && $status !== 'queued') {
                    return true;
                }
            }

            usleep(250000);
        }

        return false;
    }

    /**
     * @return array<string, mixed>
     */
    private function readJobSnapshot(string $jobsDir, string $jobId): array
    {
        $path = $jobsDir . DIRECTORY_SEPARATOR . 'servicios_marcas_' . $jobId . '.json';
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
}
