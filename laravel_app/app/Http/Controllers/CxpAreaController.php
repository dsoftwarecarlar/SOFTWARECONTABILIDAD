<?php

declare(strict_types=1);

namespace App\Http\Controllers;

use Illuminate\Contracts\View\View;

final class CxpAreaController extends Controller
{
    public function index(): View
    {
        $workspace = config('cxp.workspaces.cxp');
        $windows = array_map(
            static fn(string $windowSlug): array => config('cxp.windows.' . $windowSlug, []),
            $workspace['windows'] ?? []
        );
        $moduleCount = 0;
        foreach ($windows as $window) {
            $moduleCount += count($window['modules'] ?? []);
        }

        return view('cxp.index', [
            'workspace' => $workspace,
            'windows' => $windows,
            'moduleCount' => $moduleCount,
        ]);
    }
}
