<?php

declare(strict_types=1);

namespace App\Support;

final class AppUrl
{
    public static function asset(string $path): string
    {
        return self::path($path);
    }

    /**
     * @param array<string, mixed> $parameters
     */
    public static function route(string $name, array $parameters = []): string
    {
        return self::path(route($name, $parameters, false));
    }

    public static function path(string $path = ''): string
    {
        if (preg_match('/^https?:\/\//i', $path) === 1) {
            return $path;
        }

        $basePath = self::basePath();
        $normalized = trim($path);
        if ($normalized === '' || $normalized === '/') {
            return ($basePath !== '' ? $basePath : '') . '/';
        }

        return ($basePath !== '' ? $basePath : '') . '/' . ltrim($normalized, '/');
    }

    private static function basePath(): string
    {
        try {
            $request = request();
        } catch (\Throwable) {
            return '';
        }

        $basePath = rtrim((string) $request->getBaseUrl(), '/');
        $basePath = preg_replace('/\/index\.php$/i', '', $basePath) ?? $basePath;

        if ($basePath === '' || $basePath === '/') {
            return '';
        }

        return $basePath;
    }
}
