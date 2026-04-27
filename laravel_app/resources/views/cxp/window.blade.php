@extends('layouts.app', [
    'title' => $window['title'],
    'subtitle' => $window['description'],
    'bodyClass' => 'window-page',
])

@php
    $appUrl = \App\Support\AppUrl::class;
    $workspaceSlug = (string) ($workspace['slug'] ?? '');
    $primaryModule = $modules[0] ?? null;
@endphp

@section('hero')
    <section class="page-hero reveal-up">
        <div class="hero-grid">
            <div class="hero-copy">
                <div class="breadcrumb">
                    <a href="{{ $appUrl::route('home') }}">Inicio</a>
                    <a href="{{ $appUrl::route('workspaces.index', ['workspaceSlug' => $workspaceSlug]) }}">{{ $workspace['title'] }}</a>
                    <span>{{ $window['title'] }}</span>
                </div>
                <span class="eyebrow">Ventana operativa</span>
                <h1>{{ $window['title'] }}</h1>
                <p class="hero-note">{{ $window['description'] }}</p>
                <div class="hero-actions">
                    @if ($primaryModule)
                        <a class="primary-button" href="{{ $appUrl::route('workspaces.modules.show', ['workspaceSlug' => $workspaceSlug, 'moduleSlug' => $primaryModule['slug']]) }}">Abrir primer proceso</a>
                    @endif
                    <a class="ghost-button" href="{{ $appUrl::route('workspaces.index', ['workspaceSlug' => $workspaceSlug]) }}">Volver al area</a>
                </div>
            </div>
            <div class="hero-side">
                <article class="hero-card">
                    <span>Resumen</span>
                    <strong>{{ count($modules) }} proceso(s)</strong>
                    <p class="hero-note">
                        Aqui se concentran los procesos relacionados con esta parte del trabajo para entrar y salir sin confusion.
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
                <span class="eyebrow">Procesos</span>
                <h2>{{ $modules === [] ? 'Esta ventana ya existe y esta lista para recibir procesos.' : 'Abre el proceso exacto que necesitas.' }}</h2>
            </div>
            <p>
                {{ $modules === [] ? 'La ventana quedo registrada en la navegacion, pero todavia no tiene procesos definidos.' : 'Cada tarjeta deja claro donde entrar y para que sirve la pantalla, sin recargar al usuario con informacion innecesaria.' }}
            </p>
        </div>
        <div class="surface-grid three">
            @if ($modules === [])
                <article class="module-card">
                    <span class="chip">Sin procesos</span>
                    <h3>Ventana preparada</h3>
                    <p>Cuando se definan las acciones de esta ventana, se agregaran aqui sin rehacer la estructura general.</p>
                    <div class="runtime-badges">
                        <span class="runtime-pill">0 proceso(s)</span>
                        <span class="runtime-pill">Base lista</span>
                    </div>
                </article>
            @else
                @foreach ($modules as $module)
                    <article class="module-card">
                        <span class="chip">Proceso</span>
                        <h3>{{ $module['title'] }}</h3>
                        <p>{{ $module['mode'] === 'bundle' ? 'Consolidado final de esta ventana.' : 'Proceso disponible dentro del flujo actual.' }}</p>
                        <div class="runtime-badges">
                            <span class="runtime-pill">Disponible</span>
                            <span class="runtime-pill">{{ $module['mode'] === 'bundle' ? 'Consolidado' : 'Proceso' }}</span>
                        </div>
                        <div class="surface-actions">
                            <a class="primary-button" href="{{ $appUrl::route('workspaces.modules.show', ['workspaceSlug' => $workspaceSlug, 'moduleSlug' => $module['slug']]) }}">Abrir proceso</a>
                        </div>
                    </article>
                @endforeach
            @endif
        </div>
    </section>

    <section class="page-section reveal-up">
        <div class="section-head">
            <div>
                <span class="eyebrow">Uso rapido</span>
                <h2>La logica de esta pantalla es simple.</h2>
            </div>
            <p>
                Entra al proceso, carga los archivos del trabajo actual y descarga el resultado desde la misma pantalla.
            </p>
        </div>
        <div class="surface-grid three">
            <article class="surface-card">
                <span class="chip">1</span>
                <h3>Elegir proceso</h3>
                <p>Abre la tarjeta que corresponde al trabajo que vas a realizar.</p>
            </article>
            <article class="surface-card">
                <span class="chip">2</span>
                <h3>Cargar archivos</h3>
                <p>Sube solo los archivos solicitados por ese proceso.</p>
            </article>
            <article class="surface-card">
                <span class="chip">3</span>
                <h3>Descargar resultado</h3>
                <p>Al finalizar, revisa el archivo generado y el historial reciente.</p>
            </article>
        </div>
    </section>
@endsection
