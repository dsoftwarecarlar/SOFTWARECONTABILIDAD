<?php
declare(strict_types=1);

namespace App\Shared\Infrastructure;

final class ExternalCommandRunner
{
    /**
     * @return array{lines: list<string>, exit_code: int}
     */
    public function run(string $command): array
    {
        $lines = [];
        $exitCode = 0;

        exec($command, $lines, $exitCode);

        return [
            'lines' => array_values(array_map(static fn($line): string => (string)$line, $lines)),
            'exit_code' => $exitCode,
        ];
    }
}
