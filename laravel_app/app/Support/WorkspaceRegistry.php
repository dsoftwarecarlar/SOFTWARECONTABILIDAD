<?php

declare(strict_types=1);

namespace App\Support;

final class WorkspaceRegistry
{
    /**
     * @return array<int, array<string, mixed>>
     */
    public function workspaces(): array
    {
        $items = [];
        foreach ($this->workspaceSlugs() as $workspaceSlug) {
            $workspace = $this->workspace($workspaceSlug);
            if ($workspace !== null) {
                $items[] = $workspace;
            }
        }

        return $items;
    }

    /**
     * @return array<int, string>
     */
    public function workspaceSlugs(): array
    {
        $workspaces = config('cxp.workspaces', []);
        if (!is_array($workspaces)) {
            return [];
        }

        return array_values(array_filter(
            array_map(
                static fn ($slug): string => is_string($slug) ? trim($slug) : '',
                array_keys($workspaces)
            ),
            static fn (string $slug): bool => $slug !== ''
        ));
    }

    /**
     * @return array<string, mixed>|null
     */
    public function workspace(string $workspaceSlug): ?array
    {
        $workspace = config('cxp.workspaces.' . $workspaceSlug);
        if (!is_array($workspace)) {
            return null;
        }

        $workspace['slug'] = (string) ($workspace['slug'] ?? $workspaceSlug);

        return $workspace;
    }

    /**
     * @return array<int, array<string, mixed>>
     */
    public function windowsForWorkspace(string $workspaceSlug): array
    {
        $workspace = $this->workspace($workspaceSlug);
        if ($workspace === null) {
            return [];
        }

        $windows = [];
        foreach ((array) ($workspace['windows'] ?? []) as $windowSlug) {
            if (!is_string($windowSlug) || trim($windowSlug) === '') {
                continue;
            }

            $window = $this->windowDefinition($windowSlug);
            if ($window !== null) {
                $windows[] = $window;
            }
        }

        return $windows;
    }

    /**
     * @return array<string, mixed>|null
     */
    public function firstWindow(string $workspaceSlug): ?array
    {
        $windows = $this->windowsForWorkspace($workspaceSlug);

        return $windows[0] ?? null;
    }

    /**
     * @return array<string, mixed>|null
     */
    public function window(string $workspaceSlug, string $windowSlug): ?array
    {
        $workspace = $this->workspace($workspaceSlug);
        if ($workspace === null) {
            return null;
        }

        if (!in_array($windowSlug, (array) ($workspace['windows'] ?? []), true)) {
            return null;
        }

        return $this->windowDefinition($windowSlug);
    }

    /**
     * @return array<int, array<string, mixed>>
     */
    public function modulesForWindow(string $workspaceSlug, string $windowSlug): array
    {
        $window = $this->window($workspaceSlug, $windowSlug);
        if ($window === null) {
            return [];
        }

        $modules = [];
        foreach ((array) ($window['modules'] ?? []) as $moduleSlug) {
            if (!is_string($moduleSlug) || trim($moduleSlug) === '') {
                continue;
            }

            $module = $this->moduleDefinition($moduleSlug);
            if ($module !== null) {
                $modules[] = $module;
            }
        }

        return $modules;
    }

    /**
     * @return array<string, mixed>|null
     */
    public function moduleDefinition(string $moduleSlug): ?array
    {
        $module = config('cxp.modules.' . $moduleSlug);

        return is_array($module) ? $module : null;
    }

    /**
     * @return array{workspace: array<string, mixed>, window: array<string, mixed>|null, module: array<string, mixed>}|null
     */
    public function navigationForModule(string $workspaceSlug, string $moduleSlug): ?array
    {
        $workspace = $this->workspace($workspaceSlug);
        if ($workspace === null) {
            return null;
        }

        $module = $this->moduleDefinition($moduleSlug);
        if ($module === null) {
            return null;
        }

        foreach ($this->windowsForWorkspace($workspaceSlug) as $window) {
            if (in_array($moduleSlug, (array) ($window['modules'] ?? []), true)) {
                return [
                    'workspace' => $workspace,
                    'window' => $window,
                    'module' => $module,
                ];
            }
        }

        return null;
    }

    public function moduleCountForWorkspace(string $workspaceSlug): int
    {
        $count = 0;
        foreach ($this->windowsForWorkspace($workspaceSlug) as $window) {
            $count += count((array) ($window['modules'] ?? []));
        }

        return $count;
    }

    public function totalWindowCount(): int
    {
        $count = 0;
        foreach ($this->workspaceSlugs() as $workspaceSlug) {
            $count += count($this->windowsForWorkspace($workspaceSlug));
        }

        return $count;
    }

    public function totalModuleCount(): int
    {
        $count = 0;
        foreach ($this->workspaceSlugs() as $workspaceSlug) {
            $count += $this->moduleCountForWorkspace($workspaceSlug);
        }

        return $count;
    }

    /**
     * @return array<string, mixed>|null
     */
    private function windowDefinition(string $windowSlug): ?array
    {
        $window = config('cxp.windows.' . $windowSlug);

        return is_array($window) ? $window : null;
    }
}
