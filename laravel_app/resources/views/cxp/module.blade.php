@extends('layouts.app', [
    'title' => $module['title'],
    'subtitle' => 'Vista general del proceso',
    'bodyClass' => 'module-page',
])

@php
    $appUrl = \App\Support\AppUrl::class;
    $workspaceSlug = (string) ($workspace['slug'] ?? '');
@endphp

@section('content')
    <div class="split-layout reveal-up">
        <section class="page-section">
            <div class="breadcrumb">
                <a href="{{ $appUrl::route('home') }}">Inicio</a>
                <a href="{{ $appUrl::route('workspaces.index', ['workspaceSlug' => $workspaceSlug]) }}">{{ $workspace['title'] ?? 'Area CXP' }}</a>
                @if ($window)
                    <a href="{{ $appUrl::route('workspaces.windows.show', ['workspaceSlug' => $workspaceSlug, 'windowSlug' => $window['slug']]) }}">{{ $window['title'] }}</a>
                @endif
                <span>{{ $module['title'] }}</span>
            </div>
            <div class="section-head">
                <div>
                    <span class="eyebrow">Proceso</span>
                    <h2>{{ $module['title'] }}</h2>
                </div>
            </div>
            <p class="hero-note">Este proceso ya forma parte del portal y queda disponible desde la ventana actual.</p>
            <div class="surface-grid two" style="margin-top: 18px;">
                <article class="surface-card">
                    <span class="chip">Ubicacion</span>
                    <h3>{{ $window['title'] ?? 'Ventana actual' }}</h3>
                    <p>Usa la navegacion superior para volver a la ventana o al inicio del portal.</p>
                </article>
                <article class="surface-card">
                    <span class="chip">Estado</span>
                    <h3>Pantalla lista</h3>
                    <p>La configuracion de este proceso ya esta integrada al flujo general del area.</p>
                </article>
            </div>
        </section>

        <aside class="sidebar-stack">
            <section class="command-card">
                <span class="chip">Uso rapido</span>
                <ul class="checklist">
                    <li>Vuelve a la ventana si necesitas elegir otro proceso.</li>
                    <li>Usa la navegacion superior para cambiar de pantalla.</li>
                    <li>Trabaja siempre desde el proceso que corresponde a tu tarea.</li>
                </ul>
            </section>
        </aside>
    </div>
@endsection
