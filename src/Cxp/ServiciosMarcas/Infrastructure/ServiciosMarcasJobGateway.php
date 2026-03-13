<?php
declare(strict_types=1);

namespace App\Cxp\ServiciosMarcas\Infrastructure;

final class ServiciosMarcasJobGateway
{
    private string $jobsDir;

    public function __construct(string $jobsDir)
    {
        $this->jobsDir = $jobsDir;
    }

    public function jobsDir(): string
    {
        return $this->jobsDir;
    }

    public function read(string $jobId): ?array
    {
        if ($jobId === '' || preg_match('/^[A-Za-z0-9_-]+$/', $jobId) !== 1) {
            return null;
        }

        $path = \servicios_job_path($this->jobsDir, $jobId);
        if (!is_file($path)) {
            return null;
        }

        $data = \servicios_job_read($path);
        return $data === [] ? null : $data;
    }

    public function write(string $jobId, array $payload): void
    {
        \servicios_job_write(\servicios_job_path($this->jobsDir, $jobId), $payload);
    }

    public function refreshStaleJobs(): void
    {
        \servicios_job_refresh_stale_jobs($this->jobsDir);
    }

    /**
     * @return array<int, array<string, mixed>>
     */
    public function entries(): array
    {
        return \servicios_job_entries($this->jobsDir);
    }

    /**
     * @return string[]
     */
    public function activeStatuses(): array
    {
        return \servicios_job_active_statuses();
    }

    /**
     * @return array{count:int, job_ids: array<int, string>}
     */
    public function requestCancelAll(): array
    {
        return \servicios_job_request_cancel_all($this->jobsDir);
    }

    public function run(string $jobId, string $inputPath, string $outputDir, string $templateDir): void
    {
        \servicios_job_run($jobId, $inputPath, $outputDir, $templateDir);
    }

    /**
     * @return array<string, array{label:string, prefix:string}>
     */
    public function outputConfig(): array
    {
        return \servicios_job_output_config();
    }
}
