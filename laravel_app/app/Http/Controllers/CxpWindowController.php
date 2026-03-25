<?php

declare(strict_types=1);

namespace App\Http\Controllers;

use Illuminate\Contracts\View\View;
use Symfony\Component\HttpKernel\Exception\NotFoundHttpException;

final class CxpWindowController extends Controller
{
    public function show(string $windowSlug): View
    {
        $window = config('cxp.windows.' . $windowSlug);
        if (!is_array($window)) {
            throw new NotFoundHttpException('La ventana solicitada no existe.');
        }

        $modules = array_map(
            static fn(string $moduleSlug): array => config('cxp.modules.' . $moduleSlug, []),
            $window['modules'] ?? []
        );

        return view('cxp.window', [
            'window' => $window,
            'modules' => $modules,
            'workspace' => config('cxp.workspaces.cxp'),
        ]);
    }
}
