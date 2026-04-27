<?php

declare(strict_types=1);

namespace App\Http\Controllers;

use App\Services\CxpActionRunner;
use App\Services\RepuestosRunner;
use App\Services\ServiciosMarcasRunner;
use App\Support\WorkspaceRegistry;
use Illuminate\Http\Request;
use Illuminate\Http\JsonResponse;
use Illuminate\Http\RedirectResponse;
use Illuminate\Contracts\View\View;
use Symfony\Component\HttpKernel\Exception\NotFoundHttpException;

final class CxpModuleController extends Controller
{
    public function __construct(
        private CxpActionRunner $actions,
        private ServiciosMarcasRunner $servicios,
        private RepuestosRunner $repuestos,
        private WorkspaceRegistry $workspaces
    ) {
    }

    public function handle(Request $request, string $workspaceSlug, string $moduleSlug): View|RedirectResponse
    {
        $navigation = $this->workspaces->navigationForModule($workspaceSlug, $moduleSlug);
        if ($navigation === null) {
            throw new NotFoundHttpException('El modulo solicitado no existe.');
        }
        $module = $navigation['module'];
        $workspace = $navigation['workspace'];
        $window = $navigation['window'];

        if ($this->actions->isManagedModule($moduleSlug)) {
            $result = null;
            $error = null;
            $actionConfig = (array) ($module['action_config'] ?? config('cxp.action_modules.' . $moduleSlug, []));

            if ($request->isMethod('post')) {
                try {
                    $result = $this->actions->execute($moduleSlug, $request->file((string) ($actionConfig['upload_field'] ?? 'source_files')));
                } catch (\Throwable $exception) {
                    $error = $exception->getMessage();
                }
            }

            if ($this->actions->isBundleModule($moduleSlug)) {
                return view('cxp.bundle-module', [
                    'module' => $module,
                    'workspace' => $workspace,
                    'window' => $window,
                    'result' => $result,
                    'error' => $error,
                    'history' => $this->actions->historyFor($moduleSlug),
                    'latestActions' => $this->actions->latestActionExports(),
                    'commandPreview' => $this->commandPreview($moduleSlug),
                ]);
            }

            return view('cxp.action-module', [
                'module' => $module,
                'workspace' => $workspace,
                'window' => $window,
                'actionConfig' => $actionConfig,
                'result' => $result,
                'error' => $error,
                'history' => $this->actions->historyFor($moduleSlug),
                'commandPreview' => $this->commandPreview($moduleSlug),
            ]);
        }

        if ($this->servicios->isManagedModule($moduleSlug)) {
            $config = $this->servicios->moduleConfig();
            $this->servicios->refreshStaleJobs();
            $notice = null;
            $error = null;
            $noticeConsole = '';
            $jobId = trim((string) $request->query('job', ''));
            $job = $jobId !== '' ? $this->servicios->pendingOrResult($jobId) : null;
            $activeJobs = $this->servicios->activeJobs();
            $pendingJob = null;
            $result = null;

            $stoppedCount = $request->query('stopped');
            if ($stoppedCount !== null) {
                $count = filter_var($stoppedCount, FILTER_VALIDATE_INT, FILTER_NULL_ON_FAILURE);
                if ($count !== null) {
                    $notice = $count > 0
                        ? 'Se solicito detener ' . $count . ' proceso(s) activo(s).'
                        : 'No habia procesos activos para detener.';
                }
            }

            if ($job !== null) {
                $jobStatus = (string) ($job['status'] ?? '');
                if ($jobStatus === 'complete') {
                    $result = [
                        'source_name' => (string) ($job['source_name'] ?? ''),
                        'downloads' => is_array($job['downloads'] ?? null) ? $job['downloads'] : [],
                        'summary' => is_array($job['summary'] ?? null) ? $job['summary'] : [],
                        'console' => (string) ($job['console'] ?? ''),
                    ];
                } elseif (in_array($jobStatus, ['queued', 'running', 'cancel_requested'], true)) {
                    $pendingJob = $job;
                } elseif ($jobStatus === 'cancelled') {
                    $notice = (string) ($job['message'] ?? 'Proceso cancelado.');
                    $noticeConsole = (string) ($job['console'] ?? '');
                } elseif ($jobStatus === 'error') {
                    $error = (string) ($job['error'] ?? 'El proceso fallo.');
                }
            } elseif ($jobId !== '') {
                $notice = 'El proceso ya no esta disponible. Ejecuta uno nuevo.';
                $jobId = '';
            }

            if ($request->isMethod('post')) {
                try {
                    $action = trim((string) $request->input('action', 'process'));
                    if ($action === 'stop_all') {
                        $stopResult = $this->servicios->requestCancelAll();
                        $params = [
                            'workspaceSlug' => $workspaceSlug,
                            'moduleSlug' => $moduleSlug,
                            'stopped' => (string) $stopResult['count'],
                        ];
                        $returnJobId = trim((string) $request->input('return_job', ''));
                        if ($returnJobId !== '') {
                            $params['job'] = $returnJobId;
                        }

                        return redirect()->route('workspaces.modules.show', $params);
                    }

                    $newJobId = $this->servicios->dispatch($request->allFiles(), trim((string) $request->input('brand_key', '')));

                    return redirect()->route('workspaces.modules.show', [
                        'workspaceSlug' => $workspaceSlug,
                        'moduleSlug' => $moduleSlug,
                        'job' => $newJobId,
                    ]);
                } catch (\Throwable $exception) {
                    $error = $exception->getMessage();
                }
            }

            return view('cxp.servicios-module', [
                'module' => $module,
                'workspace' => $workspace,
                'window' => $window,
                'moduleConfig' => $config,
                'result' => $result,
                'error' => $error,
                'notice' => $notice,
                'noticeConsole' => $noticeConsole,
                'activeJobId' => $jobId,
                'activeJob' => $job,
                'pendingJob' => $pendingJob,
                'activeJobs' => $activeJobs,
                'history' => $this->servicios->history(),
                'commandPreview' => $this->servicios->commandPreview(),
                'workerCommandPreview' => $this->servicios->workerCommand(),
            ]);
        }

        if ($this->repuestos->isManagedModule($moduleSlug)) {
            $result = null;
            $error = null;
            $moduleConfig = $this->repuestos->moduleConfig();

            if ($request->isMethod('post')) {
                try {
                    $result = $this->repuestos->execute($request->allFiles());
                } catch (\Throwable $exception) {
                    $error = $exception->getMessage();
                }
            }

            return view('cxp.repuestos-module', [
                'module' => $module,
                'workspace' => $workspace,
                'window' => $window,
                'moduleConfig' => $moduleConfig,
                'result' => $result,
                'error' => $error,
                'history' => $this->repuestos->history(),
                'commandPreview' => $this->repuestos->commandPreview(),
                'workerCommandPreview' => $this->repuestos->workerCommand(),
            ]);
        }

        return view('cxp.module', [
            'module' => $module,
            'workspace' => $workspace,
            'window' => $window,
            'resourcePaths' => $this->resourcePaths($moduleSlug),
            'commandPreview' => $this->commandPreview($moduleSlug),
        ]);
    }

    public function status(string $workspaceSlug, string $moduleSlug, string $jobId): JsonResponse
    {
        if (
            !$this->servicios->isManagedModule($moduleSlug)
            || $this->workspaces->navigationForModule($workspaceSlug, $moduleSlug) === null
        ) {
            return response()->json(['status' => 'missing', 'message' => 'No se encontro el trabajo solicitado.'], 404);
        }

        $job = $this->servicios->readJob($jobId);
        if ($job === null) {
            return response()->json(['status' => 'missing', 'message' => 'No se encontro el trabajo solicitado.'], 404);
        }

        return response()->json($job);
    }

    /**
     * @return array<int, string>
     */
    private function resourcePaths(string $moduleSlug): array
    {
        return match ($moduleSlug) {
            'accion1', 'accion2', 'accion3', 'accion4', 'consolidado-acciones' => [
                (string) config('cxp.resources.acciones'),
            ],
            'servicios-marcas' => [
                (string) config('cxp.resources.servicios_marcas'),
                (string) config('cxp.storage.jobs'),
            ],
            'repuestos-tytserv' => [
                (string) config('cxp.resources.repuestos_tytserv'),
                (string) config('cxp.storage.outputs'),
                (string) config('cxp.storage.uploads'),
            ],
            default => [],
        };
    }

    /**
     * @return array{runtime: string, command: string}|null
     */
    private function commandPreview(string $moduleSlug): ?array
    {
        return match ($moduleSlug) {
            'accion1', 'accion2', 'accion3', 'accion4', 'consolidado-acciones' => $this->actions->legacyCommandFor($moduleSlug),
            'servicios-marcas' => $this->servicios->legacyCommand(),
            'repuestos-tytserv' => $this->repuestos->commandPreview(),
            default => null,
        };
    }
}
