<?php

declare(strict_types=1);

namespace App\Services;

final class ExternalProcessService
{
    /**
     * @param list<string> $arguments
     * @return array{runtime: string, command: string}
     */
    public function buildCommand(string $runtime, string $scriptPath, array $arguments = []): array
    {
        $parts = match ($runtime) {
            'node' => ['node', $scriptPath],
            'powershell+excel-com', 'powershell' => ['powershell', '-ExecutionPolicy', 'Bypass', '-File', $scriptPath],
            'python' => [config('python.binary', 'python'), $scriptPath],
            default => [$scriptPath],
        };

        foreach ($arguments as $argument) {
            $parts[] = $argument;
        }

        return [
            'runtime' => $runtime,
            'command' => implode(' ', array_map([$this, 'quote'], $parts)),
        ];
    }

    private function quote(string $value): string
    {
        if ($value === '') {
            return '""';
        }

        if (!preg_match('/[\s"]/u', $value)) {
            return $value;
        }

        return '"' . str_replace('"', '\"', $value) . '"';
    }
}

