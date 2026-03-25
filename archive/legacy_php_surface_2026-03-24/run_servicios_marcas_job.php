<?php
declare(strict_types=1);

require __DIR__ . '/includes/servicios_marcas_worker_runtime.php';

if (PHP_SAPI !== 'cli') {
    http_response_code(403);
    echo 'Solo disponible por CLI.' . PHP_EOL;
    exit(1);
}

function servicios_job_parse_args(array $argv): array
{
    $args = [];
    for ($index = 1, $count = count($argv); $index < $count; $index++) {
        $current = (string)$argv[$index];
        if (!str_starts_with($current, '--')) {
            continue;
        }

        $key = substr($current, 2);
        $value = $argv[$index + 1] ?? null;
        if (!is_string($value) || str_starts_with($value, '--')) {
            $args[$key] = '';
            continue;
        }

        $args[$key] = $value;
        $index++;
    }

    return $args;
}

$args = servicios_job_parse_args($argv);
$jobId = trim((string)($args['job'] ?? ''));
$inputPath = (string)($args['input'] ?? '');
$outputDir = (string)($args['output-dir'] ?? '');
$templateDir = (string)($args['template-dir'] ?? '');

try {
    servicios_job_run($jobId, $inputPath, $outputDir, $templateDir);
} catch (Throwable $exception) {
    fwrite(STDERR, $exception->getMessage() . PHP_EOL);
    exit(1);
}
