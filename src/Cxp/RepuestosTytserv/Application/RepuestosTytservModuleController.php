<?php
declare(strict_types=1);

namespace App\Cxp\RepuestosTytserv\Application;

use App\Cxp\RepuestosTytserv\Infrastructure\RepuestosTytservScriptGateway;

final class RepuestosTytservModuleController
{
    /**
     * @param array<string, mixed> $config
     */
    public function __construct(
        private RepuestosTytservScriptGateway $gateway,
        private array $config
    ) {
    }

    /**
     * @param array<string, mixed> $server
     * @param array<string, mixed> $files
     * @return array<string, mixed>
     */
    public function handle(array $server, array $files): array
    {
        $context = \app_workspace_module('cxp', 'cxp_repuestos_tytserv');
        if ($context === null) {
            throw new \RuntimeException('No se encontro la configuracion del modulo.');
        }

        $windowContext = \app_workspace_window('cxp', (string)($this->config['module']['window_slug'] ?? ''));
        $workspace = $context['workspace'];
        $currentModule = $context['module'];
        $window = $windowContext['window'] ?? null;
        $paths = $this->resolvePaths();
        $fileFields = $this->fileFields();
        $requestMethod = (string)($server['REQUEST_METHOD'] ?? 'GET');

        $result = null;
        $error = null;

        if ($requestMethod === 'POST') {
            try {
                if (!is_file($paths['template_path'])) {
                    throw new \RuntimeException(
                        'No existe la plantilla base ' . basename($paths['template_path']) . '.'
                    );
                }

                $stamp = date('Ymd_His');
                $savedInputs = $this->persistUploadedInputs($files, $paths['uploads_dir'], $stamp, $fileFields);
                $outputName = 'repuestos_tytserv_' . $stamp . '.xlsx';
                $outputPath = \app_join_path($paths['outputs_dir'], $outputName);

                $execution = $this->gateway->run($savedInputs, $paths['template_path'], $outputPath, $fileFields);
                $result = [
                    'excel_name' => $outputName,
                    'download_url' => \app_output_download_url($outputName),
                    'summary' => $execution['summary'],
                    'console' => $execution['console'],
                ];

                \app_cleanup_output_files_for_action('repuestos_tytserv', \app_output_retention_limit());
                \app_cleanup_upload_files(\app_upload_retention_limit());
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
            'templateFileName' => basename($paths['template_path']),
            'fileFields' => $fileFields,
            'result' => $result,
            'error' => $error,
            'history' => \app_list_output_files_for_action('repuestos_tytserv', \app_output_retention_limit()),
            'pageConfig' => $this->config['module']['page'] ?? [],
        ];
    }

    /**
     * @return array<string, string>
     */
    private function resolvePaths(): array
    {
        $templateDir = (string)($this->config['paths']['template_dir'] ?? '');
        $templateFile = (string)($this->config['module']['template_file'] ?? '');

        return [
            'uploads_dir' => \app_ensure_dir(\app_storage_path('uploads')),
            'outputs_dir' => \app_ensure_dir(\app_storage_path('outputs')),
            'template_dir' => $templateDir,
            'template_path' => \app_join_path($templateDir, $templateFile),
        ];
    }

    /**
     * @return array<int, array<string, string>>
     */
    private function fileFields(): array
    {
        return array_values(
            array_filter(
                (array)($this->config['module']['file_fields'] ?? []),
                static fn($item): bool => is_array($item)
            )
        );
    }

    /**
     * @param array<string, mixed> $files
     * @param array<int, array<string, string>> $fileFields
     * @return array<string, array{path: string}>
     */
    private function persistUploadedInputs(array $files, string $uploadsDir, string $stamp, array $fileFields): array
    {
        $savedInputs = [];
        $acceptedExtensions = $this->acceptedExtensions();

        foreach ($fileFields as $fieldConfig) {
            $field = (string)($fieldConfig['field'] ?? '');
            $label = (string)($fieldConfig['label'] ?? $field);
            if ($field === '') {
                continue;
            }

            if (!isset($files[$field])) {
                throw new \RuntimeException("No se recibio el archivo requerido: $label.");
            }

            $file = $files[$field];
            if (!is_array($file) || (int)($file['error'] ?? \UPLOAD_ERR_NO_FILE) !== \UPLOAD_ERR_OK) {
                throw new \RuntimeException("Error al subir el archivo: $label.");
            }

            $originalName = trim((string)($file['name'] ?? 'archivo.xlsx'));
            $tmpPath = (string)($file['tmp_name'] ?? '');
            $extension = strtolower(pathinfo($originalName, \PATHINFO_EXTENSION));
            if (!in_array($extension, $acceptedExtensions, true)) {
                throw new \RuntimeException("Formato no permitido en $label. Solo .xls o .xlsx.");
            }

            $safeBase = preg_replace('/[^A-Za-z0-9_-]+/', '_', pathinfo($originalName, \PATHINFO_FILENAME));
            $safeBase = trim((string)$safeBase, '_');
            if ($safeBase === '') {
                $safeBase = $field;
            }

            $inputName = sprintf('%s_%s_%s.%s', $field, $safeBase, $stamp, $extension);
            $inputPath = \app_join_path($uploadsDir, $inputName);
            if (!move_uploaded_file($tmpPath, $inputPath)) {
                throw new \RuntimeException("No se pudo guardar el archivo subido: $label.");
            }

            $savedInputs[$field] = [
                'path' => $inputPath,
            ];
        }

        return $savedInputs;
    }

    /**
     * @return string[]
     */
    private function acceptedExtensions(): array
    {
        return array_values(
            array_map(
                static fn($extension): string => strtolower((string)$extension),
                (array)($this->config['module']['accepted_extensions'] ?? ['xls', 'xlsx'])
            )
        );
    }
}
