<?php

declare(strict_types=1);

namespace App\Services;

use App\Support\AppUrl;
use Illuminate\Http\UploadedFile;
use RuntimeException;
use Symfony\Component\Process\Process;

final class CxpActionRunner
{
    private ExternalProcessService $processes;
    private ActionOutputCatalog $outputs;
    private PythonBridge $python;

    public function __construct(
        ExternalProcessService $processes,
        ActionOutputCatalog $outputs,
        PythonBridge $python
    ) {
        $this->processes = $processes;
        $this->outputs = $outputs;
        $this->python = $python;
    }

    /**
     * @return array<int, array<string, mixed>>
     */
    public function catalog(): array
    {
        return [
            config('cxp.modules.accion1'),
            config('cxp.modules.accion2'),
            config('cxp.modules.accion3'),
            config('cxp.modules.accion4'),
            config('cxp.modules.consolidado-acciones'),
        ];
    }

    /**
     * @return array<string, mixed>|null
     */
    public function moduleDefinition(string $actionSlug): ?array
    {
        $module = config('cxp.modules.' . $actionSlug);
        if (!is_array($module)) {
            return null;
        }

        $module['action_config'] = config('cxp.action_modules.' . $actionSlug, []);

        return $module;
    }

    public function isManagedModule(string $moduleSlug): bool
    {
        return is_array(config('cxp.action_modules.' . $moduleSlug));
    }

    public function isBundleModule(string $moduleSlug): bool
    {
        $definition = config('cxp.modules.' . $moduleSlug);

        return is_array($definition) && ($definition['mode'] ?? '') === 'bundle';
    }

    /**
     * @return array<int, array<string, mixed>>
     */
    public function historyFor(string $moduleSlug, int $limit = 3): array
    {
        $definition = $this->actionConfig($moduleSlug);

        return $this->outputs->listForAction((string) $definition['action_key'], $limit);
    }

    /**
     * @return array<int, array<string, mixed>>
     */
    public function latestActionExports(): array
    {
        return $this->outputs->latestActionExports();
    }

    /**
     * @param array<int, UploadedFile>|UploadedFile|null $uploaded
     * @return array<string, mixed>
     */
    public function execute(string $moduleSlug, array|UploadedFile|null $uploaded = null): array
    {
        return $this->isBundleModule($moduleSlug)
            ? $this->runBundle($moduleSlug)
            : $this->runAction($moduleSlug, $uploaded);
    }

    /**
     * @return array{runtime: string, command: string}|null
     */
    public function legacyCommandFor(string $actionSlug): ?array
    {
        $config = config('cxp.action_modules.' . $actionSlug);
        if (is_array($config) && ($config['execution_driver'] ?? '') === 'python') {
            $processor = (string) ($config['python_processor'] ?? 'probe');

            return [
                'runtime' => 'python',
                'command' => sprintf(
                    '%s %s <manifest:%s>',
                    (string) config('python.binary', 'python'),
                    base_path('../python_services/cli.py'),
                    $processor
                ),
            ];
        }

        $scripts = [
            'accion1' => base_path('../run_bot.js'),
            'accion2' => base_path('../run_bot_accion2.js'),
            'accion3' => base_path('../run_bot_accion3.js'),
            'accion4' => base_path('../run_bot_accion4.js'),
            'consolidado-acciones' => base_path('../run_export_all_actions.js'),
        ];

        if (!isset($scripts[$actionSlug])) {
            return null;
        }

        return $this->processes->buildCommand('node', $scripts[$actionSlug]);
    }

    /**
     * @param array<int, UploadedFile>|UploadedFile|null $uploaded
     * @return array<string, mixed>
     */
    private function runAction(string $moduleSlug, array|UploadedFile|null $uploaded): array
    {
        $config = $this->actionConfig($moduleSlug);
        $savedInputs = $this->persistUploadedFiles($config, $uploaded);
        $timestamp = date('Ymd_His');
        $outputBaseName = $this->resolveOutputBaseName($config, $savedInputs);
        $outputName = $outputBaseName . '_' . $timestamp . (string) $config['output_suffix'];
        $outputPath = $this->outputs->outputsDir() . DIRECTORY_SEPARATOR . $outputName;

        $execution = $this->executeActionProcess($config, $savedInputs, $outputPath);
        $generatedName = basename((string) $execution['output_path']);

        if (!is_file((string) $execution['output_path'])) {
            throw new RuntimeException('No se encontro el Excel generado.');
        }

        $this->outputs->cleanupForAction((string) $config['action_key'], 3);

        return array_filter([
            'excel_name' => $generatedName,
            'download_url' => AppUrl::route('downloads.show', ['file' => $generatedName]),
            'console' => (string) ($execution['console'] ?? ''),
            'generated_at' => date('Y-m-d H:i:s'),
            'output_origin' => (string) ($execution['output_origin'] ?? 'default_path'),
            'output_origin_note' => !empty($execution['fallback_used'])
                ? 'Se uso la ruta solicitada porque la ruta reportada por consola no fue valida.'
                : null,
            'source_count' => count($savedInputs) > 1 ? count($savedInputs) : null,
        ], static fn ($value): bool => $value !== null && $value !== '');
    }

    /**
     * @return array<string, mixed>
     */
    private function runBundle(string $moduleSlug): array
    {
        $config = $this->actionConfig($moduleSlug);
        $latestActions = $this->outputs->latestActionExports();
        $missing = array_values(array_filter(
            $latestActions,
            static fn (array $item): bool => ($item['latest'] ?? null) === null
        ));

        if ($missing !== []) {
            throw new RuntimeException(
                'Faltan archivos generados para: ' . implode(', ', array_map(
                    static fn (array $item): string => (string) $item['label'],
                    $missing
                )) . '.'
            );
        }

        $outputName = 'acciones_resumen_' . date('Ymd_His') . '.xlsx';
        $outputPath = $this->outputs->outputsDir() . DIRECTORY_SEPARATOR . $outputName;
        $execution = $this->executeBundleProcess($config, $outputPath);
        $generatedName = basename((string) $execution['output_path']);

        if (!is_file((string) $execution['output_path'])) {
            throw new RuntimeException('No se encontro el Excel consolidado generado.');
        }

        $this->outputs->cleanupForAction('bundle', 3);

        return [
            'excel_name' => $generatedName,
            'download_url' => AppUrl::route('downloads.show', ['file' => $generatedName]),
            'console' => (string) ($execution['console'] ?? ''),
            'generated_at' => date('Y-m-d H:i:s'),
        ];
    }

    /**
     * @param array<string, mixed> $config
     * @param array<int, array{original_name: string, extension: string, safe_base: string, path: string}> $savedInputs
     * @return array{output_path: string, console: string, output_origin: string, fallback_used: bool}
     */
    private function executeActionProcess(array $config, array $savedInputs, string $outputPath): array
    {
        if (($config['execution_driver'] ?? '') === 'python') {
            $result = $this->python->execute(
                (string) $config['python_processor'],
                array_map(static fn (array $item): string => $item['path'], $savedInputs),
                $outputPath,
                isset($config['template_path']) ? (string) $config['template_path'] : null,
                [
                    'script_path' => (string) $config['script_path'],
                    'command_mode' => (string) ($config['command_mode'] ?? 'positional'),
                    'generated_path_prefix' => (string) ($config['generated_path_prefix'] ?? ''),
                    'ignore_console_prefixes' => (array) ($config['ignore_console_prefixes'] ?? []),
                    'cwd' => (string) config('cxp.legacy.root'),
                ],
            );

            return [
                'output_path' => (string) $result['output_path'],
                'console' => (string) (($result['metadata']['console'] ?? '')),
                'output_origin' => (string) (($result['metadata']['output_origin'] ?? 'default_path')),
                'fallback_used' => (bool) (($result['metadata']['fallback_used'] ?? false)),
            ];
        }

        $execution = $this->runProcess($this->buildActionCommand($config, $savedInputs, $outputPath));
        $consoleLines = $this->extractConsoleLines($execution);
        $console = $this->filterConsoleLines($consoleLines, (array) ($config['ignore_console_prefixes'] ?? []));
        $artifact = $this->resolveGeneratedArtifact($outputPath, $consoleLines, (string) ($config['generated_path_prefix'] ?? ''));

        return [
            'output_path' => (string) $artifact['path'],
            'console' => $console,
            'output_origin' => (string) $artifact['origin'],
            'fallback_used' => (bool) $artifact['fallback_used'],
        ];
    }

    /**
     * @param array<string, mixed> $config
     * @return array{output_path: string, console: string, output_origin: string, fallback_used: bool}
     */
    private function executeBundleProcess(array $config, string $outputPath): array
    {
        if (($config['execution_driver'] ?? '') === 'python') {
            $result = $this->python->execute(
                (string) $config['python_processor'],
                [],
                $outputPath,
                null,
                [
                    'script_path' => (string) $config['script_path'],
                    'command_mode' => 'bundle_output_flag',
                    'generated_path_prefix' => 'Excel consolidado generado:',
                    'ignore_console_prefixes' => [],
                    'cwd' => (string) config('cxp.legacy.root'),
                ],
            );

            return [
                'output_path' => (string) $result['output_path'],
                'console' => (string) (($result['metadata']['console'] ?? '')),
                'output_origin' => (string) (($result['metadata']['output_origin'] ?? 'default_path')),
                'fallback_used' => (bool) (($result['metadata']['fallback_used'] ?? false)),
            ];
        }

        $execution = $this->runProcess([
            'node',
            (string) $config['script_path'],
            '--output',
            $outputPath,
        ]);
        $consoleLines = $this->extractConsoleLines($execution);
        $console = $this->filterConsoleLines($consoleLines, []);
        $artifact = $this->resolveGeneratedArtifact($outputPath, $consoleLines, 'Excel consolidado generado:');

        return [
            'output_path' => (string) $artifact['path'],
            'console' => $console,
            'output_origin' => (string) $artifact['origin'],
            'fallback_used' => (bool) $artifact['fallback_used'],
        ];
    }

    /**
     * @param array<int, UploadedFile>|UploadedFile|null $uploaded
     * @return array<int, array{original_name: string, extension: string, safe_base: string, path: string}>
     */
    private function persistUploadedFiles(array $config, array|UploadedFile|null $uploaded): array
    {
        $files = $this->normalizeUploadedFiles($uploaded);
        if ($files === []) {
            throw new RuntimeException('No se recibieron archivos para ejecutar la accion.');
        }

        $uploadsDir = (string) config('cxp.storage.uploads');
        if (!is_dir($uploadsDir) && !mkdir($uploadsDir, 0775, true) && !is_dir($uploadsDir)) {
            throw new RuntimeException('No se pudo preparar storage/uploads para la accion.');
        }

        $accepted = array_map('strtolower', (array) ($config['accepted_extensions'] ?? []));
        $timestamp = date('Ymd_His');
        $saved = [];

        foreach (array_values($files) as $index => $file) {
            $originalName = trim((string) $file->getClientOriginalName());
            $extension = strtolower((string) $file->getClientOriginalExtension());
            if ($accepted !== [] && !in_array($extension, $accepted, true)) {
                throw new RuntimeException('Formato no permitido. Solo se aceptan: ' . implode(', ', $accepted) . '.');
            }

            $safeBase = preg_replace('/[^A-Za-z0-9_-]+/', '_', pathinfo($originalName, PATHINFO_FILENAME));
            $safeBase = trim((string) $safeBase, '_');
            if ($safeBase === '') {
                $safeBase = 'archivo';
            }

            $inputName = count($files) > 1
                ? sprintf('%s_%s_%02d.%s', $safeBase, $timestamp, $index + 1, $extension)
                : sprintf('%s_%s.%s', $safeBase, $timestamp, $extension);

            $file->move($uploadsDir, $inputName);
            $saved[] = [
                'original_name' => $originalName,
                'extension' => $extension,
                'safe_base' => $safeBase,
                'path' => $uploadsDir . DIRECTORY_SEPARATOR . $inputName,
            ];
        }

        return $saved;
    }

    /**
     * @param array<int, UploadedFile>|UploadedFile|null $uploaded
     * @return array<int, UploadedFile>
     */
    private function normalizeUploadedFiles(array|UploadedFile|null $uploaded): array
    {
        if ($uploaded instanceof UploadedFile) {
            return [$uploaded];
        }

        if (!is_array($uploaded)) {
            return [];
        }

        return array_values(array_filter($uploaded, static fn ($item): bool => $item instanceof UploadedFile));
    }

    /**
     * @param array<int, array{original_name: string, extension: string, safe_base: string, path: string}> $savedInputs
     * @return list<string>
     */
    private function buildActionCommand(array $config, array $savedInputs, string $outputPath): array
    {
        $scriptPath = (string) $config['script_path'];
        if (!is_file($scriptPath)) {
            throw new RuntimeException('No existe el script del modulo en el proyecto.');
        }

        $command = ['node', $scriptPath];
        $commandMode = (string) ($config['command_mode'] ?? 'positional');

        return match ($commandMode) {
            'positional_action1' => array_merge($command, [
                $savedInputs[0]['path'],
                $outputPath,
            ]),
            'positional_with_template' => array_merge($command, [
                $savedInputs[0]['path'],
                $outputPath,
                (string) $config['template_path'],
            ]),
            'flagged_with_template' => array_merge(
                $command,
                array_map(static fn (array $item): string => $item['path'], $savedInputs),
                ['--output', $outputPath, '--template', (string) $config['template_path']]
            ),
            default => throw new RuntimeException('Modo de comando no soportado para el modulo.'),
        };
    }

    /**
     * @param list<string> $command
     * @return array{output: string, error: string, exit_code: int}
     */
    private function runProcess(array $command): array
    {
        $process = new Process($command, (string) config('cxp.legacy.root'));
        $process->setTimeout(180);
        $process->run();

        if (!$process->isSuccessful()) {
            $error = trim($process->getErrorOutput() . PHP_EOL . $process->getOutput());
            throw new RuntimeException($error === '' ? 'El proceso fallo al generar la salida.' : $error);
        }

        return [
            'output' => $process->getOutput(),
            'error' => $process->getErrorOutput(),
            'exit_code' => $process->getExitCode() ?? 0,
        ];
    }

    /**
     * @param array<string, mixed> $config
     * @param array<int, array{original_name: string, extension: string, safe_base: string, path: string}> $savedInputs
     */
    private function resolveOutputBaseName(array $config, array $savedInputs): string
    {
        $base = (string) ($savedInputs[0]['safe_base'] ?? 'archivo');

        return match ((string) ($config['output_base_strategy'] ?? 'first_safe_base')) {
            'first_safe_base_with_batch_suffix' => count($savedInputs) > 1 ? $base . '_lote_' . count($savedInputs) : $base,
            default => $base,
        };
    }

    /**
     * @param array{output: string, error: string, exit_code: int} $execution
     * @return list<string>
     */
    private function extractConsoleLines(array $execution): array
    {
        $combined = trim($execution['output'] . PHP_EOL . $execution['error']);
        if ($combined === '') {
            return [];
        }

        return preg_split('/\r\n|\r|\n/', $combined) ?: [];
    }

    /**
     * @param list<string> $lines
     * @param array<int, string> $ignorePrefixes
     */
    private function filterConsoleLines(array $lines, array $ignorePrefixes): string
    {
        $filtered = [];
        foreach ($lines as $line) {
            $trimmed = trim($line);
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

            if (!$ignored) {
                $filtered[] = $trimmed;
            }
        }

        return trim(implode(PHP_EOL, $filtered));
    }

    /**
     * @param list<string> $lines
     * @return array{path: string, origin: string, fallback_used: bool}
     */
    private function resolveGeneratedArtifact(string $defaultPath, array $lines, string $prefix): array
    {
        $defaultReal = realpath($defaultPath) ?: $defaultPath;
        if ($prefix === '') {
            return ['path' => $defaultReal, 'origin' => 'default_path', 'fallback_used' => false];
        }

        foreach ($lines as $line) {
            $trimmed = trim($line);
            if (stripos($trimmed, $prefix) !== 0) {
                continue;
            }

            $reported = trim(substr($trimmed, strlen($prefix)));
            $reportedReal = realpath($reported) ?: $reported;
            if ($reported !== '' && is_file($reportedReal)) {
                return ['path' => $reportedReal, 'origin' => 'console_path', 'fallback_used' => false];
            }
        }

        return ['path' => $defaultReal, 'origin' => 'default_path', 'fallback_used' => true];
    }

    /**
     * @return array<string, mixed>
     */
    private function actionConfig(string $moduleSlug): array
    {
        $config = config('cxp.action_modules.' . $moduleSlug);
        if (!is_array($config)) {
            throw new RuntimeException('No existe configuracion Laravel para este modulo.');
        }

        return $config;
    }
}
