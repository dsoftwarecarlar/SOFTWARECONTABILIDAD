@extends('layouts.app', [
    'title' => 'Portal Operativo',
    'subtitle' => config('cxp.branding.tagline'),
    'bodyClass' => 'home-page',
])

@php
    $appUrl = \App\Support\AppUrl::class;
    $branding = config('cxp.branding', []);
    $primaryLogo = !empty($branding['symbol_logo_asset']) ? $appUrl::asset($branding['symbol_logo_asset']) : null;
    $featuredWorkspace = is_array($primaryWorkspace ?? null) ? $primaryWorkspace : null;
    $featuredWindow = is_array($primaryWindow ?? null) ? $primaryWindow : null;
@endphp

@section('hero')
    <section class="page-hero home-hero reveal-up">
        <div class="hero-grid">
            <div class="hero-copy">
                <span class="eyebrow">Portal interno</span>
                <h1>Accede al area operativa y continua con el proceso del dia.</h1>
                <p class="hero-note">
                    Esta portada te lleva al punto de entrada correcto. Primero eliges el area, luego la ventana y despues el proceso.
                </p>
                <div class="hero-actions">
                    @if ($featuredWorkspace)
                        <a class="primary-button" href="{{ $appUrl::route('workspaces.index', ['workspaceSlug' => $featuredWorkspace['slug']]) }}">Entrar a {{ $featuredWorkspace['title'] }}</a>
                    @endif
                    @if ($featuredWorkspace && $featuredWindow)
                        <a class="ghost-button" href="{{ $appUrl::route('workspaces.windows.show', ['workspaceSlug' => $featuredWorkspace['slug'], 'windowSlug' => $featuredWindow['slug']]) }}">Abrir primera ventana</a>
                    @endif
                </div>
                <div class="metric-grid">
                    <article class="metric-card">
                        <span class="eyebrow">Areas</span>
                        <strong>{{ $workspaceCount }}</strong>
                        <p>La portada ya puede llevar a mas de un frente operativo sin mezclar accesos.</p>
                    </article>
                    <article class="metric-card">
                        <span class="eyebrow">Ventanas</span>
                        <strong>{{ $windowCount }}</strong>
                        <p>Las tareas se separan por tipo de trabajo para entrar mas rapido.</p>
                    </article>
                    <article class="metric-card">
                        <span class="eyebrow">Procesos</span>
                        <strong>{{ $moduleCount }}</strong>
                        <p>Cada proceso tiene su propia carga de archivos, resultado e historial.</p>
                    </article>
                </div>
            </div>
            <div class="hero-side">
                <div class="logo-showcase">
                    <div class="logo-panel">
                        @if ($primaryLogo)
                            <img src="{{ $primaryLogo }}" alt="{{ $branding['company'] ?? 'Automotores Carlos Larrea' }}" style="max-width: 220px;">
                        @endif
                    </div>
                </div>
                <article class="hero-card">
                    <span>Area destacada</span>
                    <strong>{{ $featuredWorkspace['title'] ?? ($branding['division'] ?? 'Contabilidad Talleres') }}</strong>
                    <p class="hero-note">
                        Desde aqui entras al area correcta y sigues por la ruta operativa que corresponda.
                    </p>
                </article>
            </div>
        </div>
    </section>
@endsection

@section('content')
    <section class="page-section reveal-up">
        <div class="section-head">
            <div>
                <span class="eyebrow">Como usarlo</span>
                <h2>Tres pasos para entrar y trabajar sin confusion.</h2>
            </div>
            <p>
                La prioridad es que el usuario sepa donde entrar, que archivo subir y donde descargar el resultado.
            </p>
        </div>
        <div class="surface-grid three">
            <article class="surface-card">
                <span class="chip">Paso 1</span>
                <h3>Entra al area</h3>
                <p>Desde aqui accedes al area que corresponde al trabajo que vas a realizar.</p>
            </article>
            <article class="surface-card">
                <span class="chip">Paso 2</span>
                <h3>Abre la ventana correcta</h3>
                <p>Dentro del area eliges la ventana correcta si ese frente ya tiene procesos cargados.</p>
            </article>
            <article class="surface-card">
                <span class="chip">Paso 3</span>
                <h3>Procesa y descarga</h3>
                <p>Ya dentro del proceso, cargas el archivo indicado y descargas el resultado cuando termine.</p>
            </article>
        </div>
    </section>

    <section class="page-section reveal-up">
        <div class="section-head">
            <div>
                <span class="eyebrow">Areas operativas</span>
                <h2>Elige el area con la que vas a trabajar.</h2>
            </div>
            <p>
                Cada area puede crecer con sus propias ventanas y procesos sin mezclar rutas ni responsabilidades.
            </p>
        </div>
        <div class="surface-grid one">
            @foreach ($workspaces as $workspaceItem)
                <article class="feature-panel">
                    <div>
                        <span class="chip">Area activa</span>
                        <h3>{{ $workspaceItem['title'] }}</h3>
                        <p>{{ $workspaceItem['summary'] }}</p>
                    </div>
                    <div class="surface-actions">
                        <a class="primary-button" href="{{ $appUrl::route('workspaces.index', ['workspaceSlug' => $workspaceItem['slug']]) }}">Entrar al area</a>
                    </div>
                </article>
            @endforeach
        </div>
    </section>
@endsection
