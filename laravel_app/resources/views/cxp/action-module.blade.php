@extends('layouts.app', [
    'title' => $module['title'],
    'subtitle' => $actionConfig['description'] ?? 'Proceso operativo',
    'bodyClass' => 'module-page',
])

@php
    $appUrl = \App\Support\AppUrl::class;
    $workspaceSlug = (string) ($workspace['slug'] ?? '');
    $formats = strtoupper(implode(', ', $actionConfig['accepted_extensions'] ?? []));
    $allowsMultiple = !empty($actionConfig['multiple']);
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
                <span class="eyebrow">Proceso</span>
                <h1>{{ $module['title'] }}</h1>
                <p class="hero-note">{{ $actionConfig['description'] ?? 'Carga el archivo y genera la salida correspondiente.' }}</p>
                <div class="runtime-badges">
                    <span class="runtime-pill">Formato: {{ $formats }}</span>
                    <span class="runtime-pill">{{ $allowsMultiple ? 'Carga multiple' : 'Un archivo' }}</span>
                </div>
            </div>
            <div class="hero-side">
                <article class="hero-card">
                    <span>Salida esperada</span>
                    <strong>Excel listo para descargar</strong>
                    <p class="hero-note">La pantalla esta enfocada en tres cosas: archivo correcto, resultado claro e historial reciente.</p>
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
                        <span class="eyebrow">Procesar archivo</span>
                        <h2>Carga el insumo correcto y genera la salida.</h2>
                    </div>
                    <p>Esta pantalla muestra solo lo necesario para completar la tarea y descargar el resultado.</p>
                </div>

                <form method="post" enctype="multipart/form-data" class="stack">
                    @csrf
                    <div class="upload-group">
                        <label for="source_files">{{ $actionConfig['upload_label'] }}</label>
                        <input
                            id="source_files"
                            type="file"
                            name="{{ $actionConfig['upload_field'] }}{{ $allowsMultiple ? '[]' : '' }}"
                            accept="{{ $actionConfig['accept'] }}"
                            {{ $allowsMultiple ? 'multiple' : '' }}
                            required
                        >
                        <div class="helper-text">
                            Asegurate de cargar el archivo correspondiente a este proceso.
                            @if (!empty($actionConfig['example_name']))
                                Ejemplo: <code>{{ $actionConfig['example_name'] }}</code>
                            @endif
                        </div>
                    </div>
                    <div class="button-row">
                        <button type="submit">Procesar archivo</button>
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
                        <strong>Salida generada</strong>
                        <p><code>{{ $result['excel_name'] }}</code></p>
                        @if (!empty($result['generated_at']))
                            <p>{{ $result['generated_at'] }}</p>
                        @endif
                        <div class="button-row">
                            <a class="download-link" href="{{ $result['download_url'] }}">Descargar Excel</a>
                        </div>
                    </div>
                @endif
            </section>

            <section class="page-section">
                <div class="section-head">
                    <div>
                        <span class="eyebrow">Historial</span>
                        <h2>Ultimas salidas de este proceso.</h2>
                    </div>
                </div>
                <div class="history-list">
                    @if ($history === [])
                        <article class="history-item">
                            <strong>No hay archivos generados aun.</strong>
                            <p class="meta-text">La primera salida de este proceso aparecera aqui.</p>
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
                <span class="chip">Antes de empezar</span>
                <div class="info-list">
                    <div class="info-row"><span>Formato</span><strong>{{ $formats }}</strong></div>
                    <div class="info-row"><span>Carga</span><strong>{{ $allowsMultiple ? 'Multiple' : 'Individual' }}</strong></div>
                    <div class="info-row"><span>Salida</span><strong>Excel</strong></div>
                </div>
                <ul class="checklist">
                    <li>Verifica que el archivo corresponda a este proceso.</li>
                    <li>Procesa el archivo una sola vez por carga.</li>
                    <li>Descarga el resultado desde esta misma pantalla.</li>
                </ul>
            </section>
        </aside>
    </div>
@endsection
