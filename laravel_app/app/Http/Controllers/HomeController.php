<?php

declare(strict_types=1);

namespace App\Http\Controllers;

use App\Support\WorkspaceRegistry;
use Illuminate\Contracts\View\View;

final class HomeController extends Controller
{
    public function __construct(
        private WorkspaceRegistry $workspaces
    ) {
    }

    public function index(): View
    {
        $workspaces = $this->workspaces->workspaces();
        $primaryWorkspaceSlug = (string) config('cxp.home.primary_workspace', $workspaces[0]['slug'] ?? '');
        $primaryWorkspace = $this->workspaces->workspace($primaryWorkspaceSlug);
        if ($primaryWorkspace === null) {
            $primaryWorkspace = $workspaces[0] ?? null;
        }
        $primaryWindow = is_array($primaryWorkspace ?? null)
            ? $this->workspaces->firstWindow((string) ($primaryWorkspace['slug'] ?? ''))
            : null;

        return view('home', [
            'appName' => config('app.name'),
            'primaryWorkspace' => $primaryWorkspace,
            'primaryWindow' => $primaryWindow,
            'workspaces' => $workspaces,
            'workspaceCount' => count($workspaces),
            'windowCount' => $this->workspaces->totalWindowCount(),
            'moduleCount' => $this->workspaces->totalModuleCount(),
            'pythonPlan' => config('python.planned_processors', []),
        ]);
    }
}
