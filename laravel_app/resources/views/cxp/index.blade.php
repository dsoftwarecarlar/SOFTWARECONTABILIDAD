@extends('layouts.app', [
    'title' => $workspace['title'],
    'subtitle' => $workspace['summary'],
    'bodyClass' => 'workspace-page',
])

@php
    $appUrl = \App\Support\AppUrl::class;
    $workspaceSlug = (string) ($workspace['slug'] ?? '');
    $primaryWindow = $windows[0] ?? null;
@endphp

@section('hero')
    <section class="page-hero reveal-up">
        <div class="hero-grid">
            <div class="hero-copy">
                <span class="eyebrow">Area activa</span>
                <h1>{{ $workspace['title'] }}</h1>
                <p class="hero-note">{{ $workspace['summary'] }}</p>
                <div class="hero-actions">
                    @if ($primaryWindow)
                        <a class="primary-button" href="{{ $appUrl::route('workspaces.windows.show', ['workspaceSlug' => $workspaceSlug, 'windowSlug' => $primaryWindow['slug']]) }}">Abrir primera ventana</a>
                    @endif
                    <a class="ghost-button" href="{{ $appUrl::route('home') }}">Volver al portal</a>
                </div>
                <div class="metric-grid">
                    <article class="metric-card">
                        <span class="eyebrow">Ventanas</span>
                        <strong>{{ count($windows) }}</strong>
                        <p>Entradas separadas por tipo de trabajo.</p>
                    </article>
                    <article class="metric-card">
                        <span class="eyebrow">Procesos</span>
                        <strong>{{ $moduleCount }}</strong>
                        <p>Pantallas con carga de archivos, resultado e historial.</p>
                    </article>
                    <article class="metric-card">
                        <span class="eyebrow">Estado</span>
                        <strong>{{ $windows === [] ? 'En preparacion' : 'Operativa' }}</strong>
                        <p>{{ $windows === [] ? 'La base del area ya esta creada para sumar ventanas y procesos.' : 'El equipo debe ubicar rapido cada proceso sin dudar donde entrar.' }}</p>
                    </article>
                </div>
            </div>
            <div class="hero-side">
                <article class="hero-card">
                    <span>Ruta de trabajo</span>
                    <strong>{{ $windows === [] ? 'Area lista para crecer' : 'Ventana y proceso' }}</strong>
                    <p class="hero-note">
                        {{ $windows === [] ? ($workspace['empty_description'] ?? 'Esta area todavia no tiene ventanas activas, pero ya quedo preparada para incorporarlas.') : 'Primero eliges la ventana y luego el proceso. Ese orden evita dudas y acelera el trabajo diario.' }}
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
                <span class="eyebrow">Ventanas</span>
                <h2>{{ $windows === [] ? 'Esta area ya existe y quedo lista para recibir sus procesos.' : 'Elige la ventana segun el trabajo que vas a realizar.' }}</h2>
            </div>
            <p>
                {{ $windows === [] ? ($workspace['empty_description'] ?? 'Las ventanas de esta area se iran registrando cuando se definan sus acciones.') : 'Cada ventana agrupa procesos relacionados para que el usuario vea solo lo necesario en ese momento.' }}
            </p>
        </div>
        <div class="surface-grid three">
            @if ($windows === [])
                <article class="module-card">
                    <span class="chip">Sin ventanas</span>
                    <h3>{{ $workspace['empty_title'] ?? 'Area creada.' }}</h3>
                    <p>{{ $workspace['empty_description'] ?? 'Todavia no hay ventanas registradas para esta area.' }}</p>
                    <div class="runtime-badges">
                        <span class="runtime-pill">0 proceso(s)</span>
                        <span class="runtime-pill">Base lista</span>
                    </div>
                </article>
            @else
                @foreach ($windows as $window)
                    <article class="module-card">
                        <span class="chip">Ventana</span>
                        <h3>{{ $window['title'] }}</h3>
                        <p>{{ $window['description'] }}</p>
                        <div class="runtime-badges">
                            <span class="runtime-pill">{{ count($window['modules'] ?? []) }} proceso(s)</span>
                            <span class="runtime-pill">Acceso directo</span>
                        </div>
                        <div class="surface-actions">
                            <a class="primary-button" href="{{ $appUrl::route('workspaces.windows.show', ['workspaceSlug' => $workspaceSlug, 'windowSlug' => $window['slug']]) }}">Abrir ventana</a>
                        </div>
                    </article>
                @endforeach
            @endif
        </div>
    </section>

    <section class="page-section reveal-up">
        <div class="section-head">
            <div>
                <span class="eyebrow">{{ $windows === [] ? 'Siguiente paso' : 'Antes de empezar' }}</span>
                <h2>{{ $windows === [] ? 'La estructura del area ya esta lista para cuando definamos las acciones.' : 'Que debe tener claro el usuario antes de entrar a un proceso.' }}</h2>
            </div>
            <p>
                {{ $windows === [] ? 'Cuando esta area reciba sus ventanas y procesos, ya tendra la misma navegacion limpia del resto del portal.' : 'Cada pantalla del area esta pensada para que el usuario complete una tarea y salga con un resultado claro.' }}
            </p>
        </div>
        <div class="surface-grid two">
            <article class="surface-card">
                <span class="chip">{{ $windows === [] ? 'Base' : 'Archivos' }}</span>
                <h3>{{ $windows === [] ? 'Area registrada' : 'Sube solo el insumo correcto' }}</h3>
                <p>{{ $windows === [] ? 'La nueva area ya aparece en portada, navegacion y ruta propia.' : 'Cada proceso te pide el archivo exacto que necesita y evita mostrar pasos que no hacen falta.' }}</p>
            </article>
            <article class="surface-card">
                <span class="chip">{{ $windows === [] ? 'Escala' : 'Resultado' }}</span>
                <h3>{{ $windows === [] ? 'Lista para extenderse' : 'Descarga y revisa' }}</h3>
                <p>{{ $windows === [] ? 'Las ventanas y procesos se podran agregar despues sin rehacer la estructura del portal.' : 'Al terminar, la misma pantalla deja el archivo listo para descargar y revisar en el historial reciente.' }}</p>
            </article>
        </div>
    </section>
@endsection
