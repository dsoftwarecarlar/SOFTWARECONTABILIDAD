<?php

declare(strict_types=1);

namespace App\Http\Controllers;

use App\Support\WorkspaceRegistry;
use Illuminate\Contracts\View\View;
use Symfony\Component\HttpKernel\Exception\NotFoundHttpException;

final class CxpWindowController extends Controller
{
    public function __construct(
        private WorkspaceRegistry $workspaces
    ) {
    }

    public function show(string $workspaceSlug, string $windowSlug): View
    {
        $workspace = $this->workspaces->workspace($workspaceSlug);
        $window = $this->workspaces->window($workspaceSlug, $windowSlug);
        if ($workspace === null || $window === null) {
            throw new NotFoundHttpException('La ventana solicitada no existe.');
        }

        $modules = $this->workspaces->modulesForWindow($workspaceSlug, $windowSlug);

        return view('cxp.window', [
            'window' => $window,
            'modules' => $modules,
            'workspace' => $workspace,
        ]);
    }
}
