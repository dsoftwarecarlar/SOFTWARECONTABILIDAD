<?php

declare(strict_types=1);

use Illuminate\Http\Request;

define('LARAVEL_START', microtime(true));

$autoloadPath = dirname(__DIR__) . '/vendor/autoload.php';

if (!is_file($autoloadPath)) {
    http_response_code(200);
    header('Content-Type: text/html; charset=utf-8');
    echo <<<'HTML'
<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <title>Laravel App Bootstrap</title>
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <style>
    body { font-family: "Segoe UI", sans-serif; margin: 0; background: #f5f2ea; color: #1f2933; }
    main { max-width: 880px; margin: 48px auto; padding: 0 24px; }
    .card { background: #fff; border: 1px solid #d9cfbf; border-radius: 18px; padding: 24px; box-shadow: 0 12px 40px rgba(31, 41, 51, 0.08); }
    h1 { margin-top: 0; font-size: 34px; }
    code { background: #f2ede3; padding: 2px 6px; border-radius: 6px; }
    ul { line-height: 1.6; }
  </style>
</head>
<body>
  <main>
    <div class="card">
      <h1>laravel_app lista para bootstrap</h1>
      <p>La carpeta de convivencia fue creada, pero Laravel todavia no esta instalado en este host.</p>
      <ul>
        <li>PHP actual del entorno legacy: <code>8.0.30</code></li>
        <li>Objetivo recomendado para esta app: <code>PHP 8.2+</code> con Laravel 11</li>
        <li>El sistema productivo actual no fue reemplazado</li>
        <li>Los workers activos de Node y PowerShell siguen siendo la ruta oficial</li>
      </ul>
      <p>Siguiente paso: provisionar PHP 8.2+ y Composer, luego ejecutar <code>composer install</code> dentro de <code>laravel_app</code>.</p>
    </div>
  </main>
</body>
</html>
HTML;
    exit;
}

require $autoloadPath;

$app = require_once dirname(__DIR__) . '/bootstrap/app.php';

$app->handleRequest(Request::capture());

