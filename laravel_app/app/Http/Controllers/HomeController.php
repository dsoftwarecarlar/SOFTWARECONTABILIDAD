<?php

declare(strict_types=1);

namespace App\Http\Controllers;

use Illuminate\Contracts\View\View;

final class HomeController extends Controller
{
    public function index(): View
    {
        $workspaces = array_values(config('cxp.workspaces', []));

        return view('home', [
            'appName' => config('app.name'),
            'workspace' => config('cxp.workspaces.cxp'),
            'workspaces' => $workspaces,
            'windows' => array_values(config('cxp.windows', [])),
            'workspaceCount' => count($workspaces),
            'windowCount' => count(config('cxp.windows', [])),
            'moduleCount' => count(config('cxp.modules', [])),
            'pythonPlan' => config('python.planned_processors', []),
        ]);
    }
}
