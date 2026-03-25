@extends('layouts.app', [
    'title' => $workspace['title'],
    'subtitle' => $workspace['summary'],
    'bodyClass' => 'workspace-page',
])

@php
    $appUrl = \App\Support\AppUrl::class;
@endphp

@section('hero')
    <section class="page-hero reveal-up">
        <div class="hero-grid">
            <div class="hero-copy">
                <span class="eyebrow">Area activa</span>
                <h1>{{ $workspace['title'] }}</h1>
                <p class="hero-note">{{ $workspace['summary'] }}</p>
                <div class="hero-actions">
                    <a class="primary-button" href="{{ $appUrl::route('cxp.windows.show', ['windowSlug' => 'libro-compras-aclt']) }}">Abrir libro de compras</a>
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
                        <span class="eyebrow">Objetivo</span>
                        <strong>Trabajo claro</strong>
                        <p>El equipo debe ubicar rapido cada proceso sin dudar donde entrar.</p>
                    </article>
                </div>
            </div>
            <div class="hero-side">
                <article class="hero-card">
                    <span>Ruta de trabajo</span>
                    <strong>Ventana y proceso</strong>
                    <p class="hero-note">
                        Primero eliges la ventana y luego el proceso. Ese orden evita dudas y acelera el trabajo diario.
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
                <h2>Elige la ventana segun el trabajo que vas a realizar.</h2>
            </div>
            <p>
                Cada ventana agrupa procesos relacionados para que el usuario vea solo lo necesario en ese momento.
            </p>
        </div>
        <div class="surface-grid three">
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
                        <a class="primary-button" href="{{ $appUrl::route('cxp.windows.show', ['windowSlug' => $window['slug']]) }}">Abrir ventana</a>
                    </div>
                </article>
            @endforeach
        </div>
    </section>

    <section class="page-section reveal-up">
        <div class="section-head">
            <div>
                <span class="eyebrow">Antes de empezar</span>
                <h2>Que debe tener claro el usuario antes de entrar a un proceso.</h2>
            </div>
            <p>
                Cada pantalla del area esta pensada para que el usuario complete una tarea y salga con un resultado claro.
            </p>
        </div>
        <div class="surface-grid two">
            <article class="surface-card">
                <span class="chip">Archivos</span>
                <h3>Sube solo el insumo correcto</h3>
                <p>Cada proceso te pide el archivo exacto que necesita y evita mostrar pasos que no hacen falta.</p>
            </article>
            <article class="surface-card">
                <span class="chip">Resultado</span>
                <h3>Descarga y revisa</h3>
                <p>Al terminar, la misma pantalla deja el archivo listo para descargar y revisar en el historial reciente.</p>
            </article>
        </div>
    </section>
@endsection
