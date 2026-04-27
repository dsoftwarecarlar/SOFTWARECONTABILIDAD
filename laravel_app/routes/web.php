<?php

declare(strict_types=1);

use App\Http\Controllers\CxpAreaController;
use App\Http\Controllers\CxpModuleController;
use App\Http\Controllers\CxpWindowController;
use App\Http\Controllers\DownloadController;
use App\Http\Controllers\HomeController;
use Illuminate\Support\Facades\Route;

Route::get('/', [HomeController::class, 'index'])->name('home');

$workspaceSlugs = array_values(array_filter(
    array_map(
        static fn ($slug): string => is_string($slug) ? trim($slug) : '',
        array_keys((array) config('cxp.workspaces', []))
    ),
    static fn (string $slug): bool => $slug !== ''
));

if ($workspaceSlugs !== []) {
    $workspacePattern = implode('|', array_map(
        static fn (string $slug): string => preg_quote($slug, '/'),
        $workspaceSlugs
    ));

    Route::prefix('{workspaceSlug}')
        ->where(['workspaceSlug' => $workspacePattern])
        ->name('workspaces.')
        ->group(function (): void {
            Route::get('/', [CxpAreaController::class, 'index'])->name('index');
            Route::get('/windows/{windowSlug}', [CxpWindowController::class, 'show'])->name('windows.show');
            Route::match(['get', 'post'], '/modules/{moduleSlug}', [CxpModuleController::class, 'handle'])->name('modules.show');
            Route::get('/modules/{moduleSlug}/jobs/{jobId}', [CxpModuleController::class, 'status'])->name('modules.jobs.status');
        });
}

Route::get('/downloads/{file}', [DownloadController::class, 'show'])
    ->where('file', '.*')
    ->name('downloads.show');
