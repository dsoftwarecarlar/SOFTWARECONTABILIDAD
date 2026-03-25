param(
    [string]$ListenHost = '127.0.0.1',
    [int]$Port = 8011
)

$ErrorActionPreference = 'Stop'

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..\..')
$php = Join-Path $repoRoot '.tools\php82\php.exe'
$appRoot = Join-Path $repoRoot 'laravel_app'
$router = Join-Path $appRoot 'server.php'

if (-not (Test-Path $php)) {
    throw "No existe PHP 8.2 local en $php"
}

if (-not (Test-Path $router)) {
    throw "No existe router Laravel en $router"
}

Write-Host "Laravel listo para escuchar en http://$ListenHost`:$Port/" -ForegroundColor Cyan

if ($ListenHost -eq '0.0.0.0') {
    $lanIps = Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
        Where-Object {
            $_.IPAddress -and
            $_.IPAddress -notlike '127.*' -and
            $_.IPAddress -notlike '169.254.*'
        } |
        Sort-Object InterfaceMetric, IPAddress |
        Select-Object -ExpandProperty IPAddress -Unique

    foreach ($ip in $lanIps) {
        Write-Host "Disponible por LAN: http://$ip`:$Port/" -ForegroundColor Green
    }
}

Push-Location $appRoot
try {
    & $php -d max_execution_time=180 -S "$ListenHost`:$Port" -t public $router
} finally {
    Pop-Location
}
