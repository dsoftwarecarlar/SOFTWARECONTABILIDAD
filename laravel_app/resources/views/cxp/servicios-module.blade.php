@extends('layouts.app', [
    'title' => $module['title'],
    'subtitle' => $moduleConfig['hero']['lead'] ?? 'Conciliacion mensual por marca',
    'bodyClass' => 'module-page',
])

@php
    $appUrl = \App\Support\AppUrl::class;
    $workspaceSlug = (string) ($workspace['slug'] ?? '');
    $brandGroups = [];
    foreach (($moduleConfig['upload_definitions'] ?? []) as $field => $meta) {
        $scope = (string) ($meta['scope'] ?? 'common');
        if ($scope === 'common') {
            $brandGroups['common']['label'] = 'Archivos comunes';
            $brandGroups['common']['fields'][] = ['field' => $field] + $meta;
            continue;
        }

        $brandKey = (string) ($meta['brand'] ?? 'general');
        $brandGroups[$brandKey]['label'] = strtoupper($brandKey);
        $brandGroups[$brandKey]['fields'][] = ['field' => $field] + $meta;
    }

    $statusLabels = [
        'queued' => 'Preparando archivos',
        'running' => 'Procesando',
        'cancel_requested' => 'Deteniendo',
        'complete' => 'Terminado',
        'cancelled' => 'Detenido',
        'error' => 'Con error',
    ];
@endphp

@section('hero')
    <section class="page-hero reveal-up">
        <div class="hero-grid">
            <div class="hero-copy">
                <div class="breadcrumb">
                    <a href="{{ $appUrl::route('home') }}">Inicio</a>
                    <a href="{{ $appUrl::route('workspaces.index', ['workspaceSlug' => $workspaceSlug]) }}">{{ $workspace['title'] ?? 'Area CXP' }}</a>
                    @if ($window)
                        <a href="{{ $appUrl::route('workspaces.windows.show', ['workspaceSlug' => $workspaceSlug, 'windowSlug' => $window['slug']]) }}">{{ $window['title'] }}</a>
                    @endif
                    <span>{{ $module['title'] }}</span>
                </div>
                <span class="eyebrow">{{ $moduleConfig['hero']['chip'] ?? 'Proceso mensual' }}</span>
                <h1>{{ $module['title'] }}</h1>
                <p class="hero-note">{{ $moduleConfig['hero']['lead'] ?? '' }}</p>
                <div class="runtime-badges">
                    <span class="runtime-pill">{{ count($moduleConfig['brands'] ?? []) }} marcas</span>
                    <span class="runtime-pill">2 Excel comunes</span>
                    <span class="runtime-pill">TXT por marca</span>
                </div>
            </div>
            <div class="hero-side">
                <article class="hero-card">
                    <span>Resultado esperado</span>
                    <strong>1 archivo final por marca</strong>
                    <p class="hero-note">{{ $moduleConfig['hero']['note'] ?? '' }}</p>
                </article>
            </div>
        </div>
    </section>
@endsection

@section('content')
    <div class="split-layout reveal-up">
        <div>
            @if ($notice)
                <section class="page-section">
                    <div class="feedback is-pending">
                        <strong>Aviso</strong>
                        <p>{{ $notice }}</p>
                    </div>
                </section>
            @endif

            @if ($error)
                <section class="page-section">
                    <div class="feedback is-error">
                        <strong>Error</strong>
                        <p>{{ $error }}</p>
                    </div>
                </section>
            @endif

            @if ($pendingJob)
                <section
                    class="page-section"
                    data-job-box
                    data-status-url="{{ $appUrl::route('workspaces.modules.jobs.status', ['workspaceSlug' => $workspaceSlug, 'moduleSlug' => $module['slug'], 'jobId' => $pendingJob['job_id']]) }}"
                    data-return-url="{{ $appUrl::route('workspaces.modules.show', ['workspaceSlug' => $workspaceSlug, 'moduleSlug' => $module['slug'], 'job' => $pendingJob['job_id']]) }}"
                    data-poll-interval="{{ $moduleConfig['poll']['interval_ms'] ?? 4000 }}"
                    data-poll-delay="{{ $moduleConfig['poll']['initial_delay_ms'] ?? 2000 }}"
                    data-reconnecting-message="{{ $moduleConfig['poll']['reconnecting_message'] ?? 'Actualizando estado del proceso...' }}"
                >
                    <div class="feedback is-pending">
                        <strong>Proceso activo</strong>
                        <p data-job-message>{{ $pendingJob['message'] ?? 'Procesando archivos.' }}</p>
                        <p class="meta-text" data-job-meta>
                            Estado: {{ $statusLabels[$pendingJob['status'] ?? 'running'] ?? 'Procesando' }}
                            @if (!empty($pendingJob['started_at']) || !empty($pendingJob['created_at']) || !empty($pendingJob['updated_at']))
                                | {{ $pendingJob['started_at'] ?? $pendingJob['created_at'] ?? $pendingJob['updated_at'] }}
                            @endif
                        </p>
                    </div>
                </section>
            @endif

            @if ($result)
                <section class="page-section" id="resultado">
                    <div class="section-head">
                        <div>
                            <span class="eyebrow">Salidas generadas</span>
                            <h2>Los archivos del proceso actual ya estan listos para descargar.</h2>
                        </div>
                    </div>
                    <div class="download-list">
                        @foreach ($result['downloads'] as $download)
                            <article class="result-card">
                                <div class="result-card__top">
                                    <div>
                                        <span class="chip">{{ $download['label'] }}</span>
                                        <h3>{{ $download['name'] }}</h3>
                                    </div>
                                    <a class="download-link" href="{{ $download['download_url'] }}">Descargar</a>
                                </div>
                            </article>
                        @endforeach
                    </div>

                    @if (!empty($result['summary']))
                        <div class="surface-grid two" style="margin-top: 14px;">
                            @foreach ($result['summary'] as $item)
                                <article class="surface-card">
                                    <span class="chip">{{ $item['label'] }}</span>
                                    <h3>{{ $item['rows'] }} filas procesadas</h3>
                                </article>
                            @endforeach
                        </div>
                    @endif
                </section>
            @endif

            <section class="page-section">
                <div class="section-head">
                    <div>
                        <span class="eyebrow">{{ $moduleConfig['upload_panel']['title'] ?? 'Cargar archivos' }}</span>
                        <h2>{{ $moduleConfig['upload_panel']['description'] ?? 'Carga los archivos del mes.' }}</h2>
                    </div>
                    <p>Si vas a procesar una sola marca, selecciona la marca y carga solo sus TXT. Los 2 Excel comunes siempre son obligatorios.</p>
                </div>

                <form method="post" enctype="multipart/form-data" class="stack" data-job-form>
                    @csrf
                    <input type="hidden" name="action" value="process">

                    <div class="upload-group">
                        <label for="brand_key">Marca a procesar</label>
                        <select id="brand_key" name="brand_key">
                            <option value="">Todas las marcas</option>
                            @foreach (($moduleConfig['brands'] ?? []) as $brand)
                                <option value="{{ $brand['key'] }}">{{ $brand['label'] }}</option>
                            @endforeach
                        </select>
                        <div class="helper-text">Si dejas "Todas las marcas", la pantalla exigira los TXT de cada marca.</div>
                    </div>

                    <div class="surface-grid two">
                        @foreach ($brandGroups as $brandKey => $group)
                            <section class="upload-group {{ $brandKey !== 'common' ? 'is-muted' : '' }}" @if ($brandKey !== 'common') data-brand-section="{{ $brandKey }}" @endif>
                                <span class="chip">{{ $group['label'] }}</span>
                                @foreach ($group['fields'] as $field)
                                    <div>
                                        <label for="{{ $field['field'] }}">{{ $field['label'] }}</label>
                                        <input
                                            id="{{ $field['field'] }}"
                                            type="file"
                                            name="{{ $field['field'] }}"
                                            @if (($field['scope'] ?? 'common') === 'brand') data-brand-input="{{ $field['brand'] }}" @endif
                                            accept="{{ in_array('txt', $field['accept'] ?? [], true) ? '.txt,text/plain' : '.xls,.xlsx,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }}"
                                            @if (($field['scope'] ?? 'common') === 'common') required @endif
                                        >
                                        <div class="helper-text">
                                            {{ ($field['scope'] ?? 'common') === 'common' ? 'Archivo comun obligatorio para cualquier ejecucion.' : 'Archivo requerido para la marca seleccionada.' }}
                                            @if (str_contains((string) ($field['label'] ?? ''), 'MAYOR VENTAS'))
                                                Debe corresponder al mayor de ventas y traer cuentas <code>04.01.01.xx.xxxx</code>.
                                            @endif
                                            @if (!empty($field['example_name']))
                                                Ejemplo: <code>{{ $field['example_name'] }}</code>
                                            @endif
                                        </div>
                                    </div>
                                @endforeach
                            </section>
                        @endforeach
                    </div>

                    <div class="button-row">
                        <button
                            type="submit"
                            data-submit-button
                            data-processing-label="{{ $moduleConfig['upload_panel']['processing_label'] ?? 'Procesando archivos...' }}"
                        >
                            {{ $moduleConfig['upload_panel']['button_label'] ?? 'Procesar y generar plantillas' }}
                        </button>
                    </div>
                    <p class="meta-text">{{ $moduleConfig['upload_panel']['pending_hint'] ?? '' }}</p>
                </form>
            </section>
        </div>

        <aside class="sidebar-stack">
            <section class="command-card">
                <span class="chip">Antes de empezar</span>
                <div class="info-list">
                    <div class="info-row"><span>Archivos comunes</span><strong>2</strong></div>
                    <div class="info-row"><span>Marcas</span><strong>{{ count($moduleConfig['brands'] ?? []) }}</strong></div>
                    <div class="info-row"><span>Resultado</span><strong>1 archivo por marca</strong></div>
                </div>
                <ul class="checklist">
                    <li>Ten listos los 2 Excel comunes del mes.</li>
                    <li>Verifica que los TXT correspondan a la marca correcta.</li>
                    <li>Deja esta pantalla abierta hasta que termine el proceso.</li>
                </ul>
            </section>

            @if ($activeJobs !== [])
                <section class="command-card">
                    <span class="chip">Proceso activo</span>
                    <form method="post" class="stack" data-stop-form>
                        @csrf
                        <input type="hidden" name="action" value="stop_all">
                        @if ($activeJobId !== '')
                            <input type="hidden" name="return_job" value="{{ $activeJobId }}">
                        @endif
                        <button
                            type="submit"
                            data-stop-button
                            data-confirm-message="{{ $moduleConfig['stop']['confirm_message'] ?? 'Se solicitara detener el proceso.' }}"
                            data-processing-label="{{ $moduleConfig['stop']['processing_label'] ?? 'Deteniendo proceso...' }}"
                        >
                            {{ $moduleConfig['stop']['button_label'] ?? 'Detener proceso activo' }}
                        </button>
                    </form>
                </section>
            @endif

            <section class="command-card">
                <span class="chip">{{ $moduleConfig['history_panel']['title'] ?? 'Historial Reciente' }}</span>
                <div class="history-list">
                    @if ($history === [])
                        <article class="history-item">
                            <strong>{{ $moduleConfig['history_panel']['empty_title'] ?? 'No hay archivos generados aun.' }}</strong>
                            <p class="meta-text">{{ $moduleConfig['history_panel']['empty_description'] ?? '' }}</p>
                        </article>
                    @else
                        @foreach ($history as $item)
                            <article class="history-item">
                                <div class="history-item__top">
                                    <div>
                                        <span class="chip">{{ $item['label'] }}</span>
                                        <h3>{{ $item['name'] }}</h3>
                                        <p class="meta-text">{{ $item['time'] }} | {{ number_format($item['size'] / 1024, 2) }} KB</p>
                                    </div>
                                    <a class="download-link" href="{{ $item['download_url'] }}">Descargar</a>
                                </div>
                            </article>
                        @endforeach
                    @endif
                </div>
            </section>
        </aside>
    </div>

    <script>
        (() => {
            const statusLabels = {
                queued: 'Preparando archivos',
                running: 'Procesando',
                cancel_requested: 'Deteniendo',
                complete: 'Terminado',
                cancelled: 'Detenido',
                error: 'Con error',
            };

            const form = document.querySelector('[data-job-form]');
            const submitButton = document.querySelector('[data-submit-button]');
            if (form && submitButton) {
                form.addEventListener('submit', () => {
                    submitButton.disabled = true;
                    submitButton.textContent = submitButton.getAttribute('data-processing-label') || 'Procesando archivos...';
                });
            }

            const stopForm = document.querySelector('[data-stop-form]');
            const stopButton = document.querySelector('[data-stop-button]');
            if (stopForm && stopButton && !stopButton.disabled) {
                stopForm.addEventListener('submit', (event) => {
                    const confirmMessage = stopButton.getAttribute('data-confirm-message') || 'Se solicitara detener el proceso.';
                    if (!window.confirm(confirmMessage)) {
                        event.preventDefault();
                        return;
                    }

                    stopButton.disabled = true;
                    stopButton.textContent = stopButton.getAttribute('data-processing-label') || 'Deteniendo proceso...';
                });
            }

            const brandSelect = document.querySelector('select[name="brand_key"]');
            const brandInputs = Array.from(document.querySelectorAll('[data-brand-input]'));
            const brandSections = Array.from(document.querySelectorAll('[data-brand-section]'));
            if (brandSelect && brandInputs.length > 0) {
                const syncBrandRequirements = () => {
                    const selectedBrand = brandSelect.value;
                    const requireAll = selectedBrand === '';

                    brandInputs.forEach((input) => {
                        const inputBrand = input.getAttribute('data-brand-input') || '';
                        input.required = requireAll || inputBrand === selectedBrand;
                    });

                    brandSections.forEach((section) => {
                        const sectionBrand = section.getAttribute('data-brand-section') || '';
                        section.style.opacity = (requireAll || sectionBrand === selectedBrand) ? '1' : '0.55';
                    });
                };

                brandSelect.addEventListener('change', syncBrandRequirements);
                syncBrandRequirements();
            }

            const jobBox = document.querySelector('[data-job-box]');
            if (!jobBox) {
                return;
            }

            const statusUrl = jobBox.getAttribute('data-status-url');
            const returnUrl = jobBox.getAttribute('data-return-url');
            const messageNode = jobBox.querySelector('[data-job-message]');
            const metaNode = jobBox.querySelector('[data-job-meta]');
            if (!statusUrl || !returnUrl || !messageNode || !metaNode) {
                return;
            }

            const intervalMs = Number(jobBox.getAttribute('data-poll-interval') || '4000');
            const initialDelayMs = Number(jobBox.getAttribute('data-poll-delay') || '2000');
            const reconnectingMessage = jobBox.getAttribute('data-reconnecting-message') || 'Actualizando estado del proceso...';

            const renderMeta = (job) => {
                const status = statusLabels[job.status] || 'Procesando';
                const stamp = job.started_at || job.created_at || job.updated_at || '';
                metaNode.textContent = stamp ? `Estado: ${status} | ${stamp}` : `Estado: ${status}`;
            };

            const poll = async () => {
                try {
                    const response = await fetch(statusUrl, { cache: 'no-store' });
                    if (!response.ok) {
                        throw new Error('No se pudo consultar el estado del proceso.');
                    }

                    const job = await response.json();
                    messageNode.textContent = job.message || 'Procesando archivos.';
                    renderMeta(job);

                    if (['complete', 'error', 'cancelled'].includes(job.status)) {
                        window.location.href = returnUrl + '#resultado';
                        return;
                    }
                } catch (_error) {
                    metaNode.textContent = reconnectingMessage;
                }

                window.setTimeout(poll, intervalMs);
            };

            window.setTimeout(poll, initialDelayMs);
        })();
    </script>
@endsection
