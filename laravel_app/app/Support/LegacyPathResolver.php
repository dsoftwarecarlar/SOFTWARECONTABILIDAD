<?php

declare(strict_types=1);

namespace App\Support;

final class LegacyPathResolver
{
    public function root(): string
    {
        return (string) config('cxp.legacy.root');
    }

    public function publicBaseUrl(): string
    {
        return rtrim((string) config('cxp.legacy.public_base_url'), '/');
    }

    public function absolute(string $relativePath): string
    {
        return $this->root() . DIRECTORY_SEPARATOR . ltrim(str_replace(['/', '\\'], DIRECTORY_SEPARATOR, $relativePath), DIRECTORY_SEPARATOR);
    }

    public function url(string $relativePath): string
    {
        return $this->publicBaseUrl() . '/' . ltrim(str_replace('\\', '/', $relativePath), '/');
    }

    public function legacyDownloadUrl(string $fileName): string
    {
        $downloadPath = trim((string) config('cxp.legacy.download_path', 'downloads'), '/');
        return $this->publicBaseUrl() . '/' . $downloadPath . '/' . rawurlencode($fileName);
    }
}
