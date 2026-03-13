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
        if (!is_file($this->scriptPath)) {
            throw new \RuntimeException('No existe run_repuestos_tytserv.ps1 en el proyecto.');
        }

        $commandParts = [
            escapeshellarg($this->resolvePowerShell()),
            '-NoProfile',
            '-ExecutionPolicy',
            'Bypass',
            '-File',
            escapeshellarg($this->scriptPath),
        ];

        foreach ($fileFields as $fieldConfig) {
            $field = (string)($fieldConfig['field'] ?? '');
            $flag = (string)($fieldConfig['script_flag'] ?? '');

            if ($field === '' || $flag === '' || !isset($savedInputs[$field]['path'])) {
                continue;
            }

            $commandParts[] = $flag;
            $commandParts[] = escapeshellarg($savedInputs[$field]['path']);
        }

        $commandParts[] = '-TemplatePath';
        $commandParts[] = escapeshellarg($templatePath);
        $commandParts[] = '-OutputPath';
        $commandParts[] = escapeshellarg($outputPath);

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

    private function resolvePowerShell(): string
    {
        $powershell = \app_join_path(
            getenv('WINDIR') ?: 'C:\\Windows',
            'System32',
            'WindowsPowerShell',
            'v1.0',
            'powershell.exe'
        );

        return is_file($powershell) ? $powershell : 'powershell.exe';
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
