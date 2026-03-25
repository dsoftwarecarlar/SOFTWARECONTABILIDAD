<?php
declare(strict_types=1);

namespace App\Shared\Application;

use App\Shared\Infrastructure\ExternalCommandRunner;

final class ExternalScriptModuleController
{
    /**
     * @param array<string, mixed> $config
     */
    public function __construct(
        private ExternalCommandRunner $runner,
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
        $context = \app_workspace_module($this->workspaceSlug(), $this->moduleSlug());
        if ($context === null) {
            throw new \RuntimeException('No se encontro la configuracion del modulo.');
        }

        $workspace = $context['workspace'];
        $currentModule = $context['module'];
        $uploadsDir = \app_ensure_dir(\app_storage_path('uploads'));
        $outputsDir = \app_ensure_dir(\app_storage_path('outputs'));
        $requestMethod = (string)($server['REQUEST_METHOD'] ?? 'GET');

        $result = null;
        $error = null;

        if ($requestMethod === 'POST') {
            try {
                $timestamp = date('Ymd_His');
                $savedInputs = $this->persistUploadedFiles($files, $uploadsDir, $timestamp);
                $outputBaseName = $this->resolveOutputBaseName($savedInputs, $timestamp);
                $outputFileName = $this->buildOutputFileName($outputBaseName, $timestamp, $savedInputs);
                $outputPath = \app_join_path($outputsDir, $outputFileName);

                $scriptPath = \app_join_path(\app_root(), $this->scriptRelativePath());
                if (!is_file($scriptPath)) {
                    throw new \RuntimeException($this->message('script_missing', 'No existe el script del modulo en el proyecto.'));
                }

                $execution = $this->runner->run($this->buildCommand($scriptPath, $savedInputs, $outputPath));
                $console = $this->filterConsoleLines($execution['lines']);

                if ((int)$execution['exit_code'] !== 0) {
                    throw new \RuntimeException($console === '' ? $this->message('process_failed', 'El proceso fallo.') : $console);
                }

                $generatedArtifact = $this->resolveGeneratedArtifact($outputPath, $execution['lines']);
                $generatedReal = $generatedArtifact['path'];
                $generatedName = basename($generatedReal);

                if (!is_file($generatedReal)) {
                    throw new \RuntimeException($this->message('generated_missing', 'No se encontro el Excel generado.'));
                }

                $result = [
                    'excel_name' => $generatedName,
                    'download_url' => \app_output_download_url($generatedName),
                    'console' => $console,
                    'generated_at' => date('Y-m-d H:i:s'),
                    'output_origin' => $generatedArtifact['origin'],
                ];
                if ($generatedArtifact['fallback_used']) {
                    $result['output_origin_note'] = 'Se uso ruta segura interna porque la ruta reportada por consola no fue valida.';
                }

                $result = array_merge(
                    $result,
                    $this->buildAdditionalResult($savedInputs, $execution, $generatedName)
                );

                \app_cleanup_output_files_for_action($this->actionKey(), \app_output_retention_limit());
            } catch (\Throwable $exception) {
                $error = $exception->getMessage();
            }
        }

        return [
            'brand' => \app_brand(),
            'workspace' => $workspace,
            'currentModule' => $currentModule,
            'result' => $result,
            'error' => $error,
            'history' => \app_list_output_files_for_action($this->actionKey(), \app_output_retention_limit()),
        ];
    }

    private function workspaceSlug(): string
    {
        return (string)($this->config['workspace_slug'] ?? '');
    }

    private function moduleSlug(): string
    {
        return (string)($this->config['module_slug'] ?? '');
    }

    private function actionKey(): string
    {
        return (string)($this->config['action_key'] ?? '');
    }

    private function scriptRelativePath(): string
    {
        return (string)($this->config['script_relative_path'] ?? '');
    }

    private function allowMultiple(): bool
    {
        return (bool)($this->config['allow_multiple'] ?? false);
    }

    private function uploadField(): string
    {
        return (string)($this->config['upload_field'] ?? 'file');
    }

    /**
     * @param array<string, mixed> $files
     * @return array<int, array{original_name: string, extension: string, safe_base: string, path: string}>
     */
    private function persistUploadedFiles(array $files, string $uploadsDir, string $timestamp): array
    {
        return $this->allowMultiple()
            ? $this->persistMultipleUploadedFiles($files, $uploadsDir, $timestamp)
            : [$this->persistSingleUploadedFile($files, $uploadsDir, $timestamp)];
    }

    /**
     * @param array<string, mixed> $files
     * @return array{original_name: string, extension: string, safe_base: string, path: string}
     */
    private function persistSingleUploadedFile(array $files, string $uploadsDir, string $timestamp): array
    {
        $field = $this->uploadField();
        if (!isset($files[$field])) {
            throw new \RuntimeException($this->message('missing_file', 'No se recibio el archivo requerido.'));
        }

        $file = $files[$field];
        if (!is_array($file) || (int)($file['error'] ?? \UPLOAD_ERR_NO_FILE) !== \UPLOAD_ERR_OK) {
            throw new \RuntimeException($this->message('upload_error', 'Error al subir el archivo.'));
        }

        $originalName = trim((string)($file['name'] ?? $this->defaultUploadName()));
        if ($originalName === '') {
            $originalName = $this->defaultUploadName();
        }

        $extension = $this->assertAllowedExtension($originalName);
        $safeBase = $this->safeBaseName($originalName);
        $inputFileName = $this->buildInputFileName($safeBase, $timestamp, $extension, 0, false);
        $inputPath = \app_join_path($uploadsDir, $inputFileName);
        $tmpPath = (string)($file['tmp_name'] ?? '');

        if (!move_uploaded_file($tmpPath, $inputPath)) {
            throw new \RuntimeException($this->message('save_failure', 'No se pudo guardar el archivo subido.'));
        }

        return [
            'original_name' => $originalName,
            'extension' => $extension,
            'safe_base' => $safeBase,
            'path' => $inputPath,
        ];
    }

    /**
     * @param array<string, mixed> $files
     * @return array<int, array{original_name: string, extension: string, safe_base: string, path: string}>
     */
    private function persistMultipleUploadedFiles(array $files, string $uploadsDir, string $timestamp): array
    {
        $field = $this->uploadField();
        if (!isset($files[$field])) {
            throw new \RuntimeException($this->message('missing_file', 'No se recibieron archivos.'));
        }

        $file = $files[$field];
        if (!is_array($file)) {
            throw new \RuntimeException($this->message('invalid_upload_shape', 'Formato de carga invalido.'));
        }

        $names = $file['name'] ?? [];
        $tmpNames = $file['tmp_name'] ?? [];
        $errors = $file['error'] ?? [];

        if (!is_array($names) || !is_array($tmpNames) || !is_array($errors)) {
            throw new \RuntimeException($this->message('invalid_batch_shape', 'La carga multiple no llego correctamente.'));
        }

        $savedInputs = [];

        foreach ($names as $index => $rawName) {
            $originalName = trim((string)$rawName);
            if ($originalName === '') {
                continue;
            }

            $errorCode = (int)($errors[$index] ?? \UPLOAD_ERR_NO_FILE);
            if ($errorCode === \UPLOAD_ERR_NO_FILE) {
                continue;
            }
            if ($errorCode !== \UPLOAD_ERR_OK) {
                throw new \RuntimeException($this->message('upload_error', 'Error al subir uno de los archivos.'));
            }

            $extension = $this->assertAllowedExtension($originalName);
            $safeBase = $this->safeBaseName($originalName);
            $inputFileName = $this->buildInputFileName($safeBase, $timestamp, $extension, (int)$index, true);
            $inputPath = \app_join_path($uploadsDir, $inputFileName);
            $tmpPath = (string)($tmpNames[$index] ?? '');

            if (!move_uploaded_file($tmpPath, $inputPath)) {
                throw new \RuntimeException($this->message('save_failure', 'No se pudo guardar uno de los archivos subidos.'));
            }

            $savedInputs[] = [
                'original_name' => $originalName,
                'extension' => $extension,
                'safe_base' => $safeBase,
                'path' => $inputPath,
            ];
        }

        if ($savedInputs === []) {
            throw new \RuntimeException($this->message('empty_selection', 'Selecciona al menos un archivo valido.'));
        }

        return $savedInputs;
    }

    private function defaultUploadName(): string
    {
        return (string)($this->config['default_upload_name'] ?? 'archivo');
    }

    private function fallbackBaseName(): string
    {
        return (string)($this->config['fallback_base_name'] ?? 'archivo');
    }

    private function safeBaseName(string $originalName): string
    {
        $safeBase = preg_replace('/[^A-Za-z0-9_-]+/', '_', pathinfo($originalName, \PATHINFO_FILENAME));
        $safeBase = trim((string)$safeBase, '_');

        return $safeBase === '' ? $this->fallbackBaseName() : $safeBase;
    }

    private function assertAllowedExtension(string $originalName): string
    {
        $extension = strtolower(pathinfo($originalName, \PATHINFO_EXTENSION));
        $allowed = array_values(array_map(
            static fn($value): string => strtolower((string)$value),
            (array)($this->config['accepted_extensions'] ?? [])
        ));

        if ($allowed !== [] && !in_array($extension, $allowed, true)) {
            throw new \RuntimeException($this->message('invalid_extension', 'Formato no permitido.'));
        }

        return $extension;
    }

    private function buildInputFileName(string $safeBase, string $timestamp, string $extension, int $index, bool $multiple): string
    {
        $builder = $this->config['input_name_builder'] ?? null;
        if (is_callable($builder)) {
            return (string)$builder($safeBase, $timestamp, $extension, $index, $multiple);
        }

        if ($multiple) {
            return sprintf('%s_%s_%02d.%s', $safeBase, $timestamp, $index + 1, $extension);
        }

        return sprintf('%s_%s.%s', $safeBase, $timestamp, $extension);
    }

    /**
     * @param array<int, array{original_name: string, extension: string, safe_base: string, path: string}> $savedInputs
     */
    private function resolveOutputBaseName(array $savedInputs, string $timestamp): string
    {
        $resolver = $this->config['output_base_resolver'] ?? null;
        if (is_callable($resolver)) {
            return (string)$resolver($savedInputs, $timestamp);
        }

        return (string)($savedInputs[0]['safe_base'] ?? $this->fallbackBaseName());
    }

    /**
     * @param array<int, array{original_name: string, extension: string, safe_base: string, path: string}> $savedInputs
     */
    private function buildOutputFileName(string $outputBaseName, string $timestamp, array $savedInputs): string
    {
        $builder = $this->config['output_name_builder'] ?? null;
        if (is_callable($builder)) {
            return (string)$builder($outputBaseName, $timestamp, $savedInputs);
        }

        return $outputBaseName . '_' . $timestamp . (string)($this->config['output_suffix'] ?? '_resultado.xlsx');
    }

    /**
     * @param array<int, array{original_name: string, extension: string, safe_base: string, path: string}> $savedInputs
     */
    private function buildCommand(string $scriptPath, array $savedInputs, string $outputPath): string
    {
        $builder = $this->config['command_builder'] ?? null;
        if (is_callable($builder)) {
            return (string)$builder($scriptPath, $savedInputs, $outputPath);
        }

        $parts = [
            escapeshellarg((string)($this->config['executable'] ?? 'node')),
            escapeshellarg($scriptPath),
        ];

        foreach ($savedInputs as $savedInput) {
            $parts[] = escapeshellarg($savedInput['path']);
        }

        $outputFlag = trim((string)($this->config['output_flag'] ?? ''));
        if ($outputFlag !== '') {
            $parts[] = escapeshellarg($outputFlag);
            $parts[] = escapeshellarg($outputPath);
        } elseif ((bool)($this->config['append_output_argument'] ?? true)) {
            $parts[] = escapeshellarg($outputPath);
        }

        foreach ((array)($this->config['trailing_arguments'] ?? []) as $argument) {
            $parts[] = escapeshellarg((string)$argument);
        }

        return implode(' ', $parts) . ' 2>&1';
    }

    /**
     * @param list<string> $lines
     */
    private function filterConsoleLines(array $lines): string
    {
        $consoleLines = [];
        $ignorePrefixes = array_values(array_map(
            static fn($value): string => (string)$value,
            (array)($this->config['ignore_console_prefixes'] ?? [])
        ));

        foreach ($lines as $line) {
            $trimmed = trim((string)$line);
            if ($trimmed === '') {
                continue;
            }

            $ignored = false;
            foreach ($ignorePrefixes as $prefix) {
                if ($prefix !== '' && stripos($trimmed, $prefix) === 0) {
                    $ignored = true;
                    break;
                }
            }

            if ($ignored) {
                continue;
            }

            $consoleLines[] = $trimmed;
        }

        return trim(implode(\PHP_EOL, $consoleLines));
    }

    /**
     * @param list<string> $lines
     */
    /**
     * @param list<string> $lines
     * @return array{path:string, origin:string, fallback_used:bool}
     */
    private function resolveGeneratedArtifact(string $defaultPath, array $lines): array
    {
        $defaultReal = realpath($defaultPath) ?: $defaultPath;
        $prefix = trim((string)($this->config['generated_path_prefix'] ?? ''));
        if ($prefix === '') {
            return [
                'path' => $defaultReal,
                'origin' => 'default_path',
                'fallback_used' => false,
            ];
        }

        $consolePath = '';
        foreach ($lines as $line) {
            $trimmed = trim((string)$line);
            if (stripos($trimmed, $prefix) === 0) {
                $consolePath = trim(substr($trimmed, strlen($prefix)));
                break;
            }
        }

        if ($consolePath === '') {
            return [
                'path' => $defaultReal,
                'origin' => 'default_path',
                'fallback_used' => false,
            ];
        }

        $consoleReal = realpath($consolePath);
        if ($consoleReal !== false && $this->isPathWithinOutputs($consoleReal) && is_file($consoleReal)) {
            return [
                'path' => $consoleReal,
                'origin' => 'console_path',
                'fallback_used' => false,
            ];
        }

        return [
            'path' => $defaultReal,
            'origin' => 'default_path',
            'fallback_used' => true,
        ];
    }

    private function isPathWithinOutputs(string $path): bool
    {
        $outputsDir = realpath(\app_storage_path('outputs'));
        if ($outputsDir === false) {
            return false;
        }

        $normalizedPath = str_replace(['/', '\\'], DIRECTORY_SEPARATOR, $path);
        $normalizedOutputs = rtrim(str_replace(['/', '\\'], DIRECTORY_SEPARATOR, $outputsDir), DIRECTORY_SEPARATOR);
        if (DIRECTORY_SEPARATOR === '\\') {
            $normalizedPath = strtolower($normalizedPath);
            $normalizedOutputs = strtolower($normalizedOutputs);
        }

        return $normalizedPath === $normalizedOutputs
            || str_starts_with($normalizedPath, $normalizedOutputs . DIRECTORY_SEPARATOR);
    }

    /**
     * @param array<int, array{original_name: string, extension: string, safe_base: string, path: string}> $savedInputs
     * @param array{lines: list<string>, exit_code: int} $execution
     * @return array<string, mixed>
     */
    private function buildAdditionalResult(array $savedInputs, array $execution, string $generatedName): array
    {
        $builder = $this->config['result_builder'] ?? null;
        if (!is_callable($builder)) {
            return [];
        }

        $result = $builder($savedInputs, $execution, $generatedName);
        return is_array($result) ? $result : [];
    }

    private function message(string $key, string $default): string
    {
        $messages = is_array($this->config['messages'] ?? null)
            ? $this->config['messages']
            : [];

        $value = $messages[$key] ?? $default;
        return (string)$value;
    }
}
