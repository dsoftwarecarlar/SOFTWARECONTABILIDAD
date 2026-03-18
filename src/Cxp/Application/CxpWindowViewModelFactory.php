<?php
declare(strict_types=1);

namespace App\Cxp\Application;

final class CxpWindowViewModelFactory
{
    /**
     * @param array<string, mixed> $pageConfig
     * @return array<string, mixed>
     */
    public function build(string $workspaceSlug, string $windowSlug, array $pageConfig = []): array
    {
        $context = \app_workspace_window($workspaceSlug, $windowSlug);
        if ($context === null) {
            throw new \RuntimeException('No se encontro la configuracion de la ventana.');
        }

        $window = $context['window'];

        return [
            'brand' => \app_brand(),
            'workspace' => $context['workspace'],
            'window' => $window,
            'processCount' => count($window['modules'] ?? []) + (!empty($window['bundle']) ? 1 : 0),
            'pageConfig' => $pageConfig,
            'stylesheets' => [
                \app_asset_url('css/base.css'),
                \app_asset_url('css/pages/cxp-window.css'),
            ],
        ];
    }
}
