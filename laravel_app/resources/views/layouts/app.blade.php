@php
    $appUrl = \App\Support\AppUrl::class;
    $branding = config('cxp.branding', []);
    $workspaceLinks = array_values(config('cxp.workspaces', []));
    $symbolLogo = !empty($branding['symbol_logo_asset']) ? $appUrl::asset($branding['symbol_logo_asset']) : null;
    $activeRoute = request()->route() ? request()->route()->getName() : '';
    $activeWorkspaceSlug = (string) request()->route('workspaceSlug', '');
@endphp
<!doctype html>
<html lang="es">
<head>
    <meta charset="utf-8">
    <title>{{ $title ?? config('app.name') }}</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="theme-color" content="#cb2a26">
    <link rel="stylesheet" href="{{ $appUrl::asset('assets/portal.css') }}">
</head>
<body class="portal-body {{ $bodyClass ?? '' }}">
    <div class="portal-shell">
        <header class="portal-header reveal-up">
            <a class="brand-lockup" href="{{ $appUrl::route('home') }}">
                <span class="brand-mark">
                    @if ($symbolLogo)
                        <img src="{{ $symbolLogo }}" alt="{{ $branding['company'] ?? config('app.name') }}">
                    @endif
                </span>
                <span class="brand-copy">
                    <strong>{{ $branding['company'] ?? config('app.name') }}</strong>
                    <span>{{ $branding['division'] ?? 'Operacion interna' }}</span>
                    <em>{{ $branding['platform'] ?? 'Portal' }}</em>
                </span>
            </a>

            <nav class="portal-nav" aria-label="Areas operativas">
                <a class="nav-link {{ $activeRoute === 'home' ? 'is-active' : '' }}" href="{{ $appUrl::route('home') }}">Inicio</a>
                @foreach ($workspaceLinks as $workspaceLink)
                    <a
                        class="nav-link {{ $activeWorkspaceSlug === ($workspaceLink['slug'] ?? '') ? 'is-active' : '' }}"
                        href="{{ $appUrl::route('workspaces.index', ['workspaceSlug' => $workspaceLink['slug'] ?? '']) }}"
                    >
                        {{ $workspaceLink['title'] ?? strtoupper($workspaceLink['slug'] ?? 'Area') }}
                    </a>
                @endforeach
            </nav>

            <div class="portal-status">
                <span class="status-pill is-accent">Uso interno</span>
                <span class="status-pill">{{ $branding['division'] ?? 'Contabilidad' }}</span>
            </div>
        </header>

        <main class="portal-main">
            @hasSection('hero')
                @yield('hero')
            @else
                <section class="page-hero reveal-up">
                    <div class="hero-grid">
                        <div class="hero-copy">
                            <span class="eyebrow">{{ $eyebrow ?? ($branding['platform'] ?? 'Portal Operativo') }}</span>
                            <h1>{{ $title ?? config('app.name') }}</h1>
                            @isset($subtitle)
                                <p class="hero-note">{{ $subtitle }}</p>
                            @endisset
                        </div>
                        <div class="hero-side">
                            <article class="hero-card">
                                <span>Operacion actual</span>
                                <strong>{{ $branding['division'] ?? 'Area operativa' }}</strong>
                                <p class="hero-note">{{ $branding['tagline'] ?? 'Accesos, carga de archivos e historial en un solo lugar.' }}</p>
                            </article>
                        </div>
                    </div>
                </section>
            @endif

            @yield('content')
        </main>

        <footer class="page-footer">
            {{ $branding['company'] ?? config('app.name') }} | {{ $branding['division'] ?? 'Contabilidad Talleres' }} | Accesos directos para trabajo diario.
        </footer>
    </div>
</body>
</html>
