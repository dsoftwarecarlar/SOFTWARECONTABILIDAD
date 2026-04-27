<?php

declare(strict_types=1);

namespace App\Http\Controllers;

use App\Support\WorkspaceRegistry;
use Illuminate\Contracts\View\View;
use Symfony\Component\HttpKernel\Exception\NotFoundHttpException;

final class CxpAreaController extends Controller
{
    public function __construct(
        private WorkspaceRegistry $workspaces
    ) {
    }

    public function index(string $workspaceSlug): View
    {
        $workspace = $this->workspaces->workspace($workspaceSlug);
        if ($workspace === null) {
            throw new NotFoundHttpException('El area solicitada no existe.');
        }
        $windows = $this->workspaces->windowsForWorkspace($workspaceSlug);
        $moduleCount = $this->workspaces->moduleCountForWorkspace($workspaceSlug);

        return view('cxp.index', [
            'workspace' => $workspace,
            'windows' => $windows,
            'moduleCount' => $moduleCount,
        ]);
    }
}
