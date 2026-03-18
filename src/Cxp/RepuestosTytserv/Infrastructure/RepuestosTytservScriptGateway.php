<?php
declare(strict_types=1);

namespace App\Cxp\RepuestosTytserv\Infrastructure;

use App\Shared\Infrastructure\ExternalCommandRunner;

final class RepuestosTytservScriptGateway
{
    public function __construct(
        private ExternalCommandRunner $runner,
        private string $scriptPath
    ) {
    }

    /**
     * @param array<string, array{path: string}> $savedInputs
     * @param array<int, array<string, string>> $fileFields
     * @return array{console: string, summary: array<int, array{label: string, rows: int}>}
     */
    public function run(array $savedInputs, string $templatePath, string $outputPath, array $fileFields): array
    {
        $scriptPath = $this->resolveScriptPath();
        if (!is_file($scriptPath)) {
            throw new \RuntimeException('No existe el procesador de repuestos en el proyecto.');
        }

        $isPsScript = preg_match('/\.ps1$/i', $scriptPath) === 1;
        if ($isPsScript) {
            $commandParts = [
                'powershell',
                '-NoProfile',
                '-ExecutionPolicy',
                'Bypass',
                '-File',
                escapeshellarg($scriptPath),
            ];
        } else {
            $commandParts = [
                escapeshellarg($this->resolveNodeBinary()),
                escapeshellarg($scriptPath),
            ];
        }

        foreach ($fileFields as $fieldConfig) {
            $field = (string)($fieldConfig['field'] ?? '');
            $flag = (string)($fieldConfig['script_flag'] ?? '');

            if ($field === '' || $flag === '' || !isset($savedInputs[$field]['path'])) {
                continue;
            }

            if ($isPsScript) {
                $commandParts[] = $flag;
            } else {
                $commandParts[] = $this->normalizeFlag($flag);
            }
            $commandParts[] = escapeshellarg($savedInputs[$field]['path']);
        }

        if ($isPsScript) {
            $commandParts[] = '-TemplatePath';
            $commandParts[] = escapeshellarg($templatePath);
            $commandParts[] = '-OutputPath';
            $commandParts[] = escapeshellarg($outputPath);
        } else {
            $commandParts[] = '--template-path';
            $commandParts[] = escapeshellarg($templatePath);
            $commandParts[] = '--output-path';
            $commandParts[] = escapeshellarg($outputPath);
        }

        $execution = $this->runner->run(implode(' ', $commandParts) . ' 2>&1');
        $console = trim(implode(
            \PHP_EOL,
            array_map(static fn(string $line): string => trim($line), $execution['lines'])
        ));

        if ((int)$execution['exit_code'] !== 0) {
            throw new \RuntimeException($console !== '' ? $console : 'El proceso fallo al generar la plantilla.');
        }

        if (!is_file($outputPath)) {
            throw new \RuntimeException('El proceso termino sin generar el archivo de salida.');
        }

        return [
            'console' => $console,
            'summary' => $this->parseSummary($execution['lines'], $fileFields),
        ];
    }

    private function resolveScriptPath(): string
    {
        if (is_file($this->scriptPath)) {
            return $this->scriptPath;
        }

        return \app_join_path(\app_root(), 'scripts', 'cxp', 'repuestos_tytserv', 'process.js');
    }

    private function resolveNodeBinary(): string
    {
        $programFiles = getenv('ProgramFiles') ?: 'C:\\Program Files';
        $programFilesX86 = getenv('ProgramFiles(x86)') ?: 'C:\\Program Files (x86)';
        $candidates = [
            \app_join_path($programFiles, 'nodejs', 'node.exe'),
            \app_join_path($programFilesX86, 'nodejs', 'node.exe'),
            'C:\\Program Files\\nodejs\\node.exe',
            'node',
        ];

        foreach ($candidates as $candidate) {
            if ($candidate === 'node') {
                return $candidate;
            }

            if (is_file($candidate)) {
                return $candidate;
            }
        }

        return 'node';
    }

    private function normalizeFlag(string $flag): string
    {
        return match (strtolower(trim($flag))) {
            '-inputtyt' => '--input-tyt',
            '-inputpeug' => '--input-peug',
            '-inputchgn' => '--input-chgn',
            '-inputszk' => '--input-szk',
            default => trim($flag),
        };
    }

    /**
     * @param list<string> $lines
     * @param array<int, array<string, string>> $fileFields
     * @return array<int, array{label: string, rows: int}>
     */
    private function parseSummary(array $lines, array $fileFields): array
    {
        $summary = [];

        foreach ($lines as $line) {
            $trimmed = trim($line);
            if (preg_match('/^INFO\|([a-z0-9_]+)\|rows=(\d+)$/i', $trimmed, $matches) !== 1) {
                continue;
            }

            $summary[strtolower((string)$matches[1])] = (int)$matches[2];
        }

        $orderedSummary = [];
        foreach ($fileFields as $fieldConfig) {
            $summaryKey = strtolower((string)($fieldConfig['summary_key'] ?? ''));
            if ($summaryKey === '' || !array_key_exists($summaryKey, $summary)) {
                continue;
            }

            $orderedSummary[] = [
                'label' => (string)($fieldConfig['summary_label'] ?? strtoupper($summaryKey)),
                'rows' => $summary[$summaryKey],
            ];
        }

        return $orderedSummary;
    }
}
