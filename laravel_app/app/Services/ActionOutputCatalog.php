<?php

declare(strict_types=1);

namespace App\Services;

use RuntimeException;

final class ActionOutputCatalog
{
    /**
     * @return array<string, array<string, mixed>>
     */
    public function actionExportDefinitions(): array
    {
        static $cache = null;

        if (is_array($cache)) {
            return $cache;
        }

        $path = (string) config('cxp.action_exports_config');
        if (!is_file($path)) {
            throw new RuntimeException('No existe config/cxp/action_exports.json en el proyecto.');
        }

        try {
            $decoded = json_decode((string) file_get_contents($path), true, 512, JSON_THROW_ON_ERROR);
        } catch (\Throwable $exception) {
            throw new RuntimeException('No se pudo leer config/cxp/action_exports.json: ' . $exception->getMessage(), 0, $exception);
        }

        if (!is_array($decoded)) {
            throw new RuntimeException('config/cxp/action_exports.json no contiene una lista valida de acciones.');
        }

        $definitions = [];
        foreach ($decoded as $item) {
            if (!is_array($item)) {
                continue;
            }

            $key = trim((string) ($item['key'] ?? ''));
            if ($key === '') {
                continue;
            }

            $bundleExtensions = array_values(array_filter(array_map(
                static fn ($extension): string => strtolower(ltrim(trim((string) $extension), '.')),
                is_array($item['bundle_extensions'] ?? null) ? $item['bundle_extensions'] : []
            )));

            $definitions[$key] = [
                'key' => $key,
                'label' => (string) ($item['label'] ?? $key),
                'sheet_name' => (string) ($item['sheet_name'] ?? strtoupper($key)),
                'module_path' => trim((string) ($item['module_path'] ?? ''), '/'),
                'bundle_extensions' => $bundleExtensions,
                'file_match' => is_array($item['file_match'] ?? null) ? $item['file_match'] : [],
            ];
        }

        $cache = $definitions;

        return $cache;
    }

    /**
     * @return array<int, array{name: string, path: string, size: int, timestamp: int, time: string}>
     */
    public function listForAction(string $actionKey, int $limit = 20): array
    {
        $items = array_values(array_filter(
            $this->outputFileEntries(),
            fn (array $item): bool => $this->matchesAction((string) $item['name'], $actionKey)
        ));

        return array_slice($items, 0, $limit);
    }

    /**
     * @return array<string, mixed>|null
     */
    public function latestForAction(string $actionKey): ?array
    {
        $items = $this->listForAction($actionKey, 1);

        return $items[0] ?? null;
    }

    /**
     * @return array<int, array<string, mixed>>
     */
    public function latestActionExports(): array
    {
        $actions = [];
        foreach ($this->actionExportDefinitions() as $key => $meta) {
            $actions[] = [
                'key' => $key,
                'label' => (string) $meta['label'],
                'sheet_name' => (string) $meta['sheet_name'],
                'module_slug' => $key,
                'latest' => $this->latestForAction($key),
            ];
        }

        return $actions;
    }

    public function cleanupForAction(string $actionKey, int $keep = 3): void
    {
        $entries = $this->listForAction($actionKey, 500);
        foreach (array_slice($entries, max(0, $keep)) as $entry) {
            $excelPath = (string) $entry['path'];
            if (is_file($excelPath)) {
                @unlink($excelPath);
            }

            $baseName = pathinfo((string) $entry['name'], PATHINFO_FILENAME);
            $auditPath = $this->outputsDir() . DIRECTORY_SEPARATOR . $baseName . '_auditoria.json';
            if (is_file($auditPath)) {
                @unlink($auditPath);
            }
        }
    }

    public function outputsDir(): string
    {
        return (string) config('cxp.storage.outputs');
    }

    private function matchesAction(string $fileName, string $actionKey): bool
    {
        $name = strtolower(trim($fileName));
        if (!$this->hasSupportedExtension($name)) {
            return false;
        }

        $definitions = $this->actionExportDefinitions();
        if (isset($definitions[$actionKey])) {
            return $this->matchesExportDefinition($name, $definitions[$actionKey]);
        }

        return match ($actionKey) {
            'bundle' => str_contains($name, 'acciones_resumen'),
            'servicios' => str_starts_with($name, 'servicios_'),
            'repuestos_tytserv' => str_starts_with($name, 'repuestos_tytserv_'),
            default => false,
        };
    }

    /**
     * @param array<string, mixed> $definition
     */
    private function matchesExportDefinition(string $fileName, array $definition): bool
    {
        $rule = is_array($definition['file_match'] ?? null) ? $definition['file_match'] : [];
        $type = strtolower(trim((string) ($rule['type'] ?? '')));
        $value = (string) ($rule['value'] ?? '');

        if ($type === 'contains') {
            return $value !== '' && stripos($fileName, $value) !== false;
        }

        if ($type === 'regex') {
            if ($value === '') {
                return false;
            }
            $flags = preg_replace('/[^imsxuADSUXJu]/', '', (string) ($rule['flags'] ?? '')) ?: '';
            $pattern = '~' . str_replace('~', '\\~', $value) . '~' . $flags;

            return preg_match($pattern, $fileName) === 1;
        }

        return false;
    }

    /**
     * @return array<int, array{name: string, path: string, size: int, timestamp: int, time: string}>
     */
    private function outputFileEntries(): array
    {
        $files = [];
        foreach (['xlsx', 'xls'] as $extension) {
            $pattern = $this->outputsDir() . DIRECTORY_SEPARATOR . '*.' . $extension;
            foreach (glob($pattern) ?: [] as $file) {
                if (is_file($file)) {
                    $files[] = $file;
                }
            }
        }

        $files = array_values(array_unique($files));
        usort(
            $files,
            fn (string $a, string $b): int => $this->compareFilesByTimestamp($a, $b)
        );

        $items = [];
        foreach ($files as $file) {
            if (!is_file($file)) {
                continue;
            }

            $timestamp = $this->fileTimestamp($file);
            $items[] = [
                'path' => $file,
                'name' => basename($file),
                'size' => is_file($file) ? (int) (filesize($file) ?: 0) : 0,
                'timestamp' => $timestamp,
                'time' => date('Y-m-d H:i:s', $timestamp),
            ];
        }

        return $items;
    }

    private function hasSupportedExtension(string $fileName): bool
    {
        return in_array(strtolower((string) pathinfo($fileName, PATHINFO_EXTENSION)), ['xlsx', 'xls'], true);
    }

    private function compareFilesByTimestamp(string $a, string $b): int
    {
        $timeA = $this->fileTimestamp($a);
        $timeB = $this->fileTimestamp($b);

        if ($timeA === $timeB) {
            return strcmp(basename($b), basename($a));
        }

        return $timeB <=> $timeA;
    }

    private function fileTimestamp(string $path): int
    {
        if (!is_file($path)) {
            return 0;
        }

        $created = @filectime($path);
        if (is_int($created) && $created > 0) {
            return $created;
        }

        $modified = @filemtime($path);
        if (is_int($modified) && $modified > 0) {
            return $modified;
        }

        return time();
    }
}
