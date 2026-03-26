<?php

declare(strict_types=1);

namespace App\Services;

use App\Support\AppUrl;
use Illuminate\Http\UploadedFile;
use RuntimeException;

final class RepuestosRunner
{
    public function __construct(
        private ExternalProcessService $processes,
        private ActionOutputCatalog $outputs,
        private PythonBridge $python
    ) {
    }

    public function isManagedModule(string $moduleSlug): bool
    {
        return $moduleSlug === 'repuestos-tytserv';
    }

    /**
     * @return array<string, mixed>
     */
    public function moduleConfig(): array
    {
        $config = config('cxp.repuestos_module');
        if (!is_array($config)) {
            throw new RuntimeException('No existe configuracion Laravel para Repuestos TYTSERV.');
        }

        return $config;
    }

    /**
     * @return array<int, array<string, mixed>>
     */
    public function history(int $limit = 3): array
    {
        return $this->outputs->listForAction('repuestos_tytserv', $limit);
    }

    /**
     * @param array<string, UploadedFile|array<int, UploadedFile>|null> $uploadedFiles
     * @return array<string, mixed>
     */
    public function execute(array $uploadedFiles): array
    {
        $config = $this->moduleConfig();
        $templatePath = (string) ($config['template_path'] ?? '');
        if (!is_file($templatePath)) {
            throw new RuntimeException('No existe la plantilla base de Repuestos TYTSERV.');
        }

        $stamp = date('Ymd_His') . '_' . $this->nonce();
        $savedInputs = $this->persistUploadedInputs($uploadedFiles, $stamp, $config);
        $outputName = 'repuestos_tytserv_' . $stamp . '.xlsx';
        $outputPath = $this->outputs->outputsDir() . DIRECTORY_SEPARATOR . $outputName;

        $result = $this->python->execute(
            (string) ($config['python_processor'] ?? 'repuestos_tytserv.process'),
            array_map(static fn (array $item): string => $item['path'], $savedInputs),
            $outputPath,
            $templatePath,
            [
                'script_path' => (string) ($config['script_path'] ?? ''),
                'file_fields' => (array) ($config['file_fields'] ?? []),
                'saved_inputs' => $savedInputs,
                'cwd' => (string) config('cxp.legacy.root'),
            ]
        );

        $generatedPath = (string) ($result['output_path'] ?? $outputPath);
        if (!is_file($generatedPath)) {
            throw new RuntimeException('El proceso termino sin generar el archivo de salida.');
        }

        $keepOutputs = max(1, (int) ($config['output_retention_limit'] ?? 3));
        $this->outputs->cleanupForAction('repuestos_tytserv', $keepOutputs);
        $this->cleanupUploads(max(1, (int) ($config['upload_retention_limit'] ?? 20)));

        return [
            'excel_name' => basename($generatedPath),
            'download_url' => AppUrl::route('downloads.show', ['file' => basename($generatedPath)]),
            'summary' => is_array($result['metadata']['summary'] ?? null) ? $result['metadata']['summary'] : [],
            'integrity_checks' => is_array($result['metadata']['integrity_checks'] ?? null) ? $result['metadata']['integrity_checks'] : [],
            'console' => (string) ($result['metadata']['console'] ?? ''),
            'generated_at' => date('Y-m-d H:i:s'),
            'output_origin' => (string) ($result['metadata']['output_origin'] ?? 'default_path'),
        ];
    }

    /**
     * @return array{runtime: string, command: string}
     */
    public function commandPreview(): array
    {
        $config = $this->moduleConfig();

        return [
            'runtime' => 'python',
            'command' => sprintf(
                '%s %s <manifest:%s>',
                (string) config('python.binary', 'python'),
                base_path('../python_services/cli.py'),
                (string) ($config['python_processor'] ?? 'repuestos_tytserv.process')
            ),
        ];
    }

    /**
     * @return array{runtime: string, command: string}
     */
    public function workerCommand(): array
    {
        $config = $this->moduleConfig();

        return $this->processes->buildCommand(
            'node',
            (string) ($config['script_path'] ?? base_path('../scripts/cxp/repuestos_tytserv/process.js'))
        );
    }

    private function nonce(): string
    {
        try {
            return bin2hex(random_bytes(3));
        } catch (\Throwable) {
            return substr(sha1((string) microtime(true) . (string) mt_rand()), 0, 6);
        }
    }

    /**
     * @param array<string, UploadedFile|array<int, UploadedFile>|null> $uploadedFiles
     * @param array<string, mixed> $config
     * @return array<string, array{field: string, path: string, original_name: string, brand_key: string, source_label: string}>
     */
    private function persistUploadedInputs(array $uploadedFiles, string $stamp, array $config): array
    {
        $uploadsDir = (string) config('cxp.storage.uploads');
        if (!is_dir($uploadsDir) && !mkdir($uploadsDir, 0775, true) && !is_dir($uploadsDir)) {
            throw new RuntimeException('No se pudo preparar storage/uploads para Repuestos TYTSERV.');
        }

        $accepted = array_map(
            static fn ($extension): string => strtolower((string) $extension),
            (array) ($config['accepted_extensions'] ?? ['xls', 'xlsx'])
        );

        $savedInputs = [];
        foreach ((array) ($config['file_fields'] ?? []) as $fieldConfig) {
            if (!is_array($fieldConfig)) {
                continue;
            }

            $field = (string) ($fieldConfig['field'] ?? '');
            $label = (string) ($fieldConfig['label'] ?? $field);
            if ($field === '') {
                continue;
            }

            $uploaded = $uploadedFiles[$field] ?? null;
            if (!$uploaded instanceof UploadedFile) {
                throw new RuntimeException("No se recibio el archivo requerido: {$label}.");
            }

            $originalName = trim((string) $uploaded->getClientOriginalName());
            $extension = strtolower((string) $uploaded->getClientOriginalExtension());
            if ($accepted !== [] && !in_array($extension, $accepted, true)) {
                throw new RuntimeException("Formato no permitido en {$label}. Solo .xls o .xlsx.");
            }

            $safeBase = preg_replace('/[^A-Za-z0-9_-]+/', '_', pathinfo($originalName, PATHINFO_FILENAME));
            $safeBase = trim((string) $safeBase, '_');
            if ($safeBase === '') {
                $safeBase = $field;
            }

            $inputName = sprintf('%s_%s_%s.%s', $field, $safeBase, $stamp, $extension);
            $uploaded->move($uploadsDir, $inputName);
            $savedInputs[$field] = [
                'field' => $field,
                'path' => $uploadsDir . DIRECTORY_SEPARATOR . $inputName,
                'original_name' => $originalName,
                'brand_key' => (string) ($fieldConfig['brand_key'] ?? ''),
                'source_label' => (string) ($fieldConfig['source_label'] ?? ''),
            ];
        }

        return $savedInputs;
    }

    private function cleanupUploads(int $keep): void
    {
        $uploadsDir = (string) config('cxp.storage.uploads');
        if (!is_dir($uploadsDir)) {
            return;
        }

        $files = glob($uploadsDir . DIRECTORY_SEPARATOR . '*.xls*') ?: [];
        usort(
            $files,
            static function (string $a, string $b): int {
                $timeA = filemtime($a) ?: 0;
                $timeB = filemtime($b) ?: 0;
                if ($timeA === $timeB) {
                    return strcmp(basename($b), basename($a));
                }

                return $timeB <=> $timeA;
            }
        );

        foreach (array_slice($files, $keep) as $file) {
            @unlink($file);
        }
    }
}
