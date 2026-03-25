<?php
declare(strict_types=1);

namespace App\Cxp\ServiciosMarcas\Application;

final class ServiciosMarcasWindowViewModelFactory
{
    /**
     * @param array<string, mixed> $config
     * @return array<string, mixed>
     */
    public function build(array $config): array
    {
        $context = \app_workspace_window('cxp', 'conciliacion_servicios_marcas');
        if ($context === null) {
            throw new \RuntimeException('No se encontro la configuracion de la ventana.');
        }

        $window = $context['window'];

        return [
            'brand' => \app_brand(),
            'workspace' => $context['workspace'],
            'window' => $window,
            'processCount' => count($window['modules'] ?? []),
            'pageConfig' => $config['window'] ?? [],
            'stylesheets' => [
                \app_asset_url('css/base.css'),
                \app_asset_url('css/pages/cxp-window-servicios-marcas.css'),
            ],
        ];
    }
}
