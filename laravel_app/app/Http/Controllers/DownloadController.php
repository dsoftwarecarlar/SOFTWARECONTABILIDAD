<?php

declare(strict_types=1);

namespace App\Http\Controllers;

use App\Support\LegacyPathResolver;
use Illuminate\Http\RedirectResponse;
use Symfony\Component\HttpFoundation\BinaryFileResponse;

final class DownloadController extends Controller
{
    private LegacyPathResolver $paths;

    public function __construct(
        LegacyPathResolver $paths
    ) {
        $this->paths = $paths;
    }

    public function show(string $file): BinaryFileResponse|RedirectResponse
    {
        $outputsDir = realpath((string) config('cxp.storage.outputs', ''));
        $candidate = $outputsDir
            ? realpath($outputsDir . DIRECTORY_SEPARATOR . basename(str_replace(['/', '\\'], DIRECTORY_SEPARATOR, $file)))
            : false;

        if (is_string($candidate) && is_file($candidate) && $this->isPathWithinDirectory($candidate, $outputsDir ?: '')) {
            return response()->download($candidate, basename($candidate));
        }

        return redirect()->away($this->paths->legacyDownloadUrl($file));
    }

    private function isPathWithinDirectory(string $path, string $directory): bool
    {
        if ($directory === '') {
            return false;
        }

        $normalizedPath = str_replace(['/', '\\'], DIRECTORY_SEPARATOR, $path);
        $normalizedDirectory = rtrim(str_replace(['/', '\\'], DIRECTORY_SEPARATOR, $directory), DIRECTORY_SEPARATOR);

        if (DIRECTORY_SEPARATOR === '\\') {
            $normalizedPath = strtolower($normalizedPath);
            $normalizedDirectory = strtolower($normalizedDirectory);
        }

        return $normalizedPath === $normalizedDirectory
            || str_starts_with($normalizedPath, $normalizedDirectory . DIRECTORY_SEPARATOR);
    }
}
