@extends('layouts.app', [
    'title' => $module['title'],
    'subtitle' => $moduleConfig['page']['hero_lead'] ?? 'Proceso guiado para facturacion mensual de repuestos',
    'bodyClass' => 'module-page',
])

@php
    $appUrl = \App\Support\AppUrl::class;
    $brandGroups = [];
    foreach (($moduleConfig['file_fields'] ?? []) as $fieldConfig) {
        $brandKey = (string) ($fieldConfig['brand_key'] ?? ($fieldConfig['field'] ?? 'general'));
        if (!isset($brandGroups[$brandKey])) {
            $brandGroups[$brandKey] = [
                'label' => (string) ($fieldConfig['brand_label'] ?? strtoupper($brandKey)),
                'fields' => [],
            ];
        }

        $brandGroups[$brandKey]['fields'][] = $fieldConfig;
    }
@endphp

@section('hero')
    <section class="page-hero reveal-up">
        <div class="hero-grid">
            <div class="hero-copy">
                <div class="breadcrumb">
                    <a href="{{ $appUrl::route('home') }}">Inicio</a>
                    <a href="{{ $appUrl::route('cxp.index') }}">{{ $workspace['title'] ?? 'Area CXP' }}</a>
                    @if ($window)
                        <a href="{{ $appUrl::route('cxp.windows.show', ['windowSlug' => $window['slug']]) }}">{{ $window['title'] }}</a>
                    @endif
                    <span>{{ $module['title'] }}</span>
                </div>
                <span class="eyebrow">{{ $moduleConfig['page']['hero_chip'] ?? 'Proceso mensual' }}</span>
                <h1>{{ $module['title'] }}</h1>
                <p class="hero-note">{{ $moduleConfig['page']['hero_lead'] ?? '' }}</p>
                <div class="runtime-badges">
                    <span class="runtime-pill">{{ count($brandGroups) }} marcas</span>
                    <span class="runtime-pill">{{ count($moduleConfig['file_fields'] ?? []) }} archivos del mes</span>
                </div>
            </div>
            <div class="hero-side">
                <article class="hero-card">
                    <span>Resultado esperado</span>
                    <strong>1 libro final con todas las hojas del proceso</strong>
                    <p class="hero-note">{{ $moduleConfig['page']['hero_note'] ?? '' }}</p>
                </article>
            </div>
        </div>
    </section>
@endsection

@section('content')
    <div class="split-layout reveal-up">
        <div>
            <section class="page-section">
                <div class="section-head">
                    <div>
                        <span class="eyebrow">{{ $moduleConfig['page']['process_title'] ?? 'Cargar archivos del mes' }}</span>
                        <h2>{{ $moduleConfig['page']['process_description'] ?? 'Carga los archivos requeridos y genera el reporte final.' }}</h2>
                    </div>
                    <p>Organiza primero los Excel de ventas y devoluciones de cada marca. Luego procesa todo en una sola ejecucion.</p>
                </div>

                <form method="post" enctype="multipart/form-data" class="stack">
                    @csrf
                    <div class="surface-grid two">
                        @foreach ($brandGroups as $group)
                            <section class="upload-group">
                                <span class="chip">{{ $group['label'] }}</span>
                                @foreach ($group['fields'] as $fieldConfig)
                                    <div>
                                        <label for="{{ $fieldConfig['field'] }}">{{ $fieldConfig['source_label'] }}</label>
                                        <input
                                            id="{{ $fieldConfig['field'] }}"
                                            type="file"
                                            name="{{ $fieldConfig['field'] }}"
                                            accept=".xls,.xlsx,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                            required
                                        >
                                        <div class="helper-text">
                                            Carga el Excel correspondiente a {{ strtolower($fieldConfig['source_label']) }} de {{ $group['label'] }}.
                                            @if (!empty($fieldConfig['example_name']))
                                                Ejemplo: <code>{{ $fieldConfig['example_name'] }}</code>
                                            @endif
                                            @if (($fieldConfig['source_label'] ?? '') === 'Devoluciones')
                                                Si origen exporta devoluciones vacias como un .xlsx sin hojas, tambien se acepta.
                                            @endif
                                        </div>
                                    </div>
                                @endforeach
                            </section>
                        @endforeach
                    </div>

                    <div class="button-row">
                        <button type="submit">{{ $moduleConfig['page']['button_label'] ?? 'Procesar y generar reporte' }}</button>
                    </div>
                </form>

                @if ($error)
                    <div class="feedback is-error">
                        <strong>Error</strong>
                        <p>{{ $error }}</p>
                    </div>
                @endif

                @if ($result)
                    <div class="feedback is-success" id="resultado">
                        <strong>Archivo listo</strong>
                        <p><code>{{ $result['excel_name'] }}</code></p>
                        @if (!empty($result['generated_at']))
                            <p>{{ $result['generated_at'] }}</p>
                        @endif
                        <div class="button-row">
                            <a class="download-link" href="{{ $result['download_url'] }}">Descargar reporte final</a>
                        </div>
                    </div>

                    @if (!empty($result['summary']))
                        <div class="surface-grid two">
                            @foreach ($result['summary'] as $item)
                                <article class="surface-card">
                                    <span class="chip">{{ $item['label'] }}</span>
                                    <h3>{{ $item['rows'] }} filas procesadas</h3>
                                </article>
                            @endforeach
                        </div>
                    @endif

                    @if (!empty($result['integrity_checks']))
                        <div class="surface-grid two">
                            @foreach ($result['integrity_checks'] as $stage => $status)
                                <article class="surface-card">
                                    <span class="chip">{{ strtoupper(str_replace('_', ' ', $stage)) }}</span>
                                    <h3>{{ strtoupper((string) $status) }}</h3>
                                    <p class="meta-text">Verificacion interna del libro final.</p>
                                </article>
                            @endforeach
                        </div>
                    @endif
                @endif
            </section>

            <section class="page-section">
                <div class="section-head">
                    <div>
                        <span class="eyebrow">{{ $moduleConfig['page']['history_title'] ?? 'Historial reciente' }}</span>
                        <h2>Archivos recientes del proceso.</h2>
                    </div>
                </div>
                <div class="history-list">
                    @if ($history === [])
                        <article class="history-item">
                            <strong>{{ $moduleConfig['page']['history_empty_title'] ?? 'No hay reportes generados aun.' }}</strong>
                            <p class="meta-text">{{ $moduleConfig['page']['history_empty_description'] ?? '' }}</p>
                        </article>
                    @else
                        @foreach ($history as $item)
                            <article class="history-item">
                                <div class="history-item__top">
                                    <div>
                                        <strong>{{ $item['name'] }}</strong>
                                        <p class="meta-text">{{ $item['time'] }} | {{ number_format($item['size'] / 1024, 2) }} KB</p>
                                    </div>
                                    <a class="download-link" href="{{ $appUrl::route('downloads.show', ['file' => $item['name']]) }}">Descargar</a>
                                </div>
                            </article>
                        @endforeach
                    @endif
                </div>
            </section>
        </div>

        <aside class="sidebar-stack">
            <section class="command-card">
                <span class="chip">{{ $moduleConfig['page']['inputs_title'] ?? 'Antes de empezar' }}</span>
                <div class="info-list">
                    <div class="info-row"><span>Marcas</span><strong>{{ count($brandGroups) }}</strong></div>
                    <div class="info-row"><span>Archivos requeridos</span><strong>{{ count($moduleConfig['file_fields'] ?? []) }}</strong></div>
                    <div class="info-row"><span>Resultado</span><strong>1 libro final</strong></div>
                </div>
                <ul class="checklist">
                    <li>Ten listos los 8 Excel del mes antes de iniciar.</li>
                    <li>Revisa que cada archivo corresponda a la marca correcta.</li>
                    <li>Descarga el libro final cuando termine el proceso.</li>
                </ul>
                <p class="meta-text">{{ $moduleConfig['page']['retention_description'] ?? '' }}</p>
            </section>
        </aside>
    </div>
@endsection
