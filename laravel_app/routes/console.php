<?php

declare(strict_types=1);

use App\Services\PythonBridge;
use Illuminate\Support\Facades\Artisan;

Artisan::command('cxp:about', function (): void {
    $this->comment('Laravel coexistence scaffold listo. El sistema legacy sigue siendo la ruta activa.');
})->purpose('Describe el estado de la migracion coexistente.');

Artisan::command('python:probe {manifest?}', function (PythonBridge $python, ?string $manifest = null): void {
    $result = $python->probe($manifest);

    $this->line('Command: ' . implode(' ', $result['command']));
    $this->line('Exit code: ' . (string) $result['exit_code']);
    $this->newLine();
    $this->line($result['output']);
})->purpose('Verifica que Laravel puede invocar la capa Python futura.');
