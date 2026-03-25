<?php

declare(strict_types=1);

use App\Http\Controllers\CxpAreaController;
use App\Http\Controllers\CxpModuleController;
use App\Http\Controllers\CxpWindowController;
use App\Http\Controllers\DownloadController;
use App\Http\Controllers\HomeController;
use Illuminate\Support\Facades\Route;

Route::get('/', [HomeController::class, 'index'])->name('home');

Route::prefix('cxp')->name('cxp.')->group(function (): void {
    Route::get('/', [CxpAreaController::class, 'index'])->name('index');
    Route::get('/windows/{windowSlug}', [CxpWindowController::class, 'show'])->name('windows.show');
    Route::match(['get', 'post'], '/modules/{moduleSlug}', [CxpModuleController::class, 'handle'])->name('modules.show');
    Route::get('/modules/{moduleSlug}/jobs/{jobId}', [CxpModuleController::class, 'status'])->name('modules.jobs.status');
});

Route::get('/downloads/{file}', [DownloadController::class, 'show'])
    ->where('file', '.*')
    ->name('downloads.show');
