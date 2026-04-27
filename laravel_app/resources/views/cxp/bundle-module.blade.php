@extends('layouts.app', [
    'title' => $module['title'],
    'subtitle' => 'Consolidado de la ventana',
    'bodyClass' => 'module-page',
])

@php
    $appUrl = \App\Support\AppUrl::class;
    $workspaceSlug = (string) ($workspace['slug'] ?? '');
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
                <span class="eyebrow">Consolidado</span>
                <h1>{{ $module['title'] }}</h1>
                <p class="hero-note">Reune en un solo archivo las ultimas salidas disponibles de esta ventana.</p>
            </div>
            <div class="hero-side">
                <article class="hero-card">
                    <span>Uso sugerido</span>
                    <strong>Revisar y unificar</strong>
                    <p class="hero-note">Primero valida que cada proceso tenga su salida lista y luego genera el consolidado final.</p>
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
                        <span class="eyebrow">Cobertura</span>
                        <h2>Ultima salida disponible por proceso.</h2>
                    </div>
                    <p>Antes de construir el consolidado, confirma que cada proceso ya tenga una salida reciente.</p>
                </div>
                <div class="surface-grid two">
                    @foreach ($latestActions as $action)
                        <article class="surface-card">
                            <span class="chip">{{ $action['label'] }}</span>
                            @if ($action['latest'])
                                <h3>{{ $action['latest']['name'] }}</h3>
                                <p>{{ $action['latest']['time'] }}</p>
                            @else
                                <h3>Sin salida reciente</h3>
                                <p>Falta generar este archivo antes del consolidado.</p>
                            @endif
                        </article>
                    @endforeach
                </div>
            </section>

            <section class="page-section">
                <div class="section-head">
                    <div>
                        <span class="eyebrow">Generar</span>
                        <h2>Construir el archivo consolidado.</h2>
                    </div>
                </div>
                <form method="post" class="stack">
                    @csrf
                    <div class="button-row">
                        <button type="submit">Construir consolidado</button>
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
                            <a class="download-link" href="{{ $result['download_url'] }}">Descargar consolidado</a>
                        </div>
                    </div>
                @endif
            </section>
        </div>

        <aside class="sidebar-stack">
            <section class="command-card">
                <span class="chip">Antes de generar</span>
                <ul class="checklist">
                    <li>Revisa que cada proceso tenga su archivo listo.</li>
                    <li>Genera el consolidado una vez finalizadas las acciones del mes.</li>
                    <li>Descarga el archivo final desde esta misma pantalla.</li>
                </ul>
            </section>

            <section class="command-card">
                <span class="chip">Historial</span>
                <div class="history-list">
                    @if ($history === [])
                        <article class="history-item">
                            <strong>No hay consolidados generados aun.</strong>
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
        </aside>
    </div>
@endsection
