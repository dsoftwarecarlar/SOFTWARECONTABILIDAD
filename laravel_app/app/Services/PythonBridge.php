<?php

declare(strict_types=1);

namespace App\Services;

use RuntimeException;
use Symfony\Component\Process\Process;

final class PythonBridge
{
    /**
     * @return array{command: list<string>, output: string, exit_code: int}
     */
    public function probe(?string $manifestPath = null): array
    {
        $pythonBinary = (string) config('python.binary', 'python');
        $root = rtrim((string) config('python.root'), DIRECTORY_SEPARATOR);
        $manifest = $manifestPath ?: $root . DIRECTORY_SEPARATOR . 'samples' . DIRECTORY_SEPARATOR . 'request.example.json';
        $cli = $root . DIRECTORY_SEPARATOR . 'cli.py';

        $command = [$pythonBinary, $cli, $manifest];
        $process = new Process($command, base_path('..'));
        $process->setTimeout(30);
        $process->run();

        return [
            'command' => $command,
            'output' => trim($process->getOutput() . PHP_EOL . $process->getErrorOutput()),
            'exit_code' => $process->getExitCode() ?? 1,
        ];
    }

    /**
     * @param list<string> $inputPaths
     * @return array{success: bool, label: string, output_path: string, metadata: array<string, mixed>}
     */
    public function execute(
        string $processor,
        array $inputPaths,
        string $outputPath,
        ?string $templatePath = null,
        array $options = []
    ): array {
        $payload = [
            'processor' => $processor,
            'input_paths' => $inputPaths,
            'output_path' => $outputPath,
            'template_path' => $templatePath,
            'options' => $options,
        ];

        $manifestsDir = storage_path('app/python_manifests');
        if (!is_dir($manifestsDir) && !mkdir($manifestsDir, 0775, true) && !is_dir($manifestsDir)) {
            throw new RuntimeException('No se pudo preparar storage/app/python_manifests para el bridge Python.');
        }

        $manifestPath = $manifestsDir . DIRECTORY_SEPARATOR . $processor . '_' . date('Ymd_His') . '_' . bin2hex(random_bytes(4)) . '.json';
        $encoded = json_encode($payload, JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES | JSON_THROW_ON_ERROR);
        file_put_contents($manifestPath, $encoded);

        $result = $this->runManifest($manifestPath);
        if ($result['exit_code'] !== 0) {
            $message = $result['error'] ?? trim($result['output'] . PHP_EOL . $result['raw_output']);
            throw new RuntimeException($message === '' ? 'El bridge Python fallo al ejecutar el procesador.' : $message);
        }

        return $result['payload'] ?? throw new RuntimeException('El bridge Python no devolvio un payload valido.');
    }

    /**
     * @return array{
     *   command: list<string>,
     *   output: string,
     *   raw_output: string,
     *   exit_code: int,
     *   payload?: array{success: bool, label: string, output_path: string, metadata: array<string, mixed>},
     *   error?: string
     * }
     */
    private function runManifest(string $manifestPath): array
    {
        $pythonBinary = (string) config('python.binary', 'python');
        $root = rtrim((string) config('python.root'), DIRECTORY_SEPARATOR);
        $cli = $root . DIRECTORY_SEPARATOR . 'cli.py';

        $command = [$pythonBinary, $cli, $manifestPath];
        $process = new Process($command, base_path('..'));
        $process->setTimeout(300);
        $process->run();

        $rawOutput = trim($process->getOutput() . PHP_EOL . $process->getErrorOutput());
        $decoded = null;
        if ($rawOutput !== '') {
            try {
                $decoded = json_decode($rawOutput, true, 512, JSON_THROW_ON_ERROR);
            } catch (\Throwable) {
                $decoded = null;
            }
        }

        return [
            'command' => $command,
            'output' => is_array($decoded) ? json_encode($decoded, JSON_UNESCAPED_SLASHES | JSON_THROW_ON_ERROR) : $rawOutput,
            'raw_output' => $rawOutput,
            'exit_code' => $process->getExitCode() ?? 1,
            'payload' => is_array($decoded) ? $decoded : null,
            'error' => is_array($decoded) ? (string) ($decoded['error'] ?? '') : null,
        ];
    }
}
