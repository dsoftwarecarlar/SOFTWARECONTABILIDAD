@extends('layouts.app', [
    'title' => 'Portal Operativo',
    'subtitle' => config('cxp.branding.tagline'),
    'bodyClass' => 'home-page',
])

@php
    $appUrl = \App\Support\AppUrl::class;
    $branding = config('cxp.branding', []);
    $primaryLogo = !empty($branding['symbol_logo_asset']) ? $appUrl::asset($branding['symbol_logo_asset']) : null;
@endphp

@section('hero')
    <section class="page-hero home-hero reveal-up">
        <div class="hero-grid">
            <div class="hero-copy">
                <span class="eyebrow">Portal interno</span>
                <h1>Accede al area operativa y continua con el proceso del dia.</h1>
                <p class="hero-note">
                    Esta portada te lleva al punto de entrada correcto. Desde el area eliges la ventana y luego el proceso.
                </p>
                <div class="hero-actions">
                    <a class="primary-button" href="{{ $appUrl::route('cxp.index') }}">Entrar a {{ $workspace['title'] }}</a>
                    <a class="ghost-button" href="{{ $appUrl::route('cxp.windows.show', ['windowSlug' => 'libro-compras-aclt']) }}">Abrir libro de compras</a>
                </div>
                <div class="metric-grid">
                    <article class="metric-card">
                        <span class="eyebrow">Area</span>
                        <strong>{{ $workspace['title'] }}</strong>
                        <p>Todo el trabajo actual de esta portada entra por una sola area operativa.</p>
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
                    <span>Area activa</span>
                    <strong>{{ $branding['division'] ?? 'Contabilidad Talleres' }}</strong>
                    <p class="hero-note">
                        Ingresa al area y abre la ventana que corresponde al trabajo que vas a realizar.
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
                <p>Desde aqui accedes al espacio de trabajo principal de contabilidad talleres.</p>
            </article>
            <article class="surface-card">
                <span class="chip">Paso 2</span>
                <h3>Abre la ventana correcta</h3>
                <p>Dentro del area eliges si vas a trabajar libro de compras, servicios por marca o repuestos.</p>
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
                <span class="eyebrow">Entrada principal</span>
                <h2>Contabilidad Talleres entra por una sola puerta.</h2>
            </div>
            <p>
                Si necesitas volver al punto de partida, entra al area y desde ahi continua con la ventana y el proceso correspondiente.
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
                        <a class="primary-button" href="{{ $appUrl::route(($workspaceItem['slug'] ?? 'cxp') . '.index') }}">Entrar al area</a>
                    </div>
                </article>
            @endforeach
        </div>
    </section>
@endsection
