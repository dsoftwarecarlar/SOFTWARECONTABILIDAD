<?php
declare(strict_types=1);

namespace App\Cxp\Application;

final class CxpWorkspaceHomeViewModelFactory
{
    /**
     * @param array<string, mixed> $pageConfig
     * @return array<string, mixed>
     */
    public function build(string $workspaceSlug, array $pageConfig = []): array
    {
        $workspace = \app_workspace_by_slug($workspaceSlug);
        if ($workspace === null) {
            throw new \RuntimeException('No se encontro la configuracion del area.');
        }

        return [
            'brand' => \app_brand(),
            'workspace' => $workspace,
            'windowCount' => count($workspace['windows'] ?? []),
            'publicRootUrl' => \app_public_root_url(),
            'pageConfig' => $pageConfig,
            'stylesheets' => [
                \app_asset_url('css/base.css'),
                \app_asset_url('css/pages/cxp-workspace.css'),
            ],
        ];
    }
}
