param(
    [int]$Port = 8085,
    [switch]$TryOpenFirewall
)

$ErrorActionPreference = 'Stop'

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..\..')
$php = Join-Path $repoRoot '.tools\php82\php.exe'
$appRoot = Join-Path $repoRoot 'laravel_app'
$router = Join-Path $appRoot 'server.php'
$logDir = Join-Path $appRoot 'storage\logs'
$outLog = Join-Path $logDir "laravel_lan_$Port.out.log"
$errLog = Join-Path $logDir "laravel_lan_$Port.err.log"
$pidFile = Join-Path $logDir "laravel_lan_$Port.pid"

if (-not (Test-Path $php)) {
    throw "No existe PHP 8.2 local en $php"
}

if (-not (Test-Path $router)) {
    throw "No existe router Laravel en $router"
}

New-Item -ItemType Directory -Path $logDir -Force | Out-Null

if (Test-Path $pidFile) {
    $existingPid = (Get-Content $pidFile -ErrorAction SilentlyContinue | Select-Object -First 1)
    if ($existingPid) {
        $existing = Get-Process -Id $existingPid -ErrorAction SilentlyContinue
        if ($existing) {
            throw "Ya existe un servidor Laravel LAN activo con PID $existingPid en el puerto $Port."
        }
    }
    Remove-Item $pidFile -Force -ErrorAction SilentlyContinue
}

$arguments = @(
    '-d', 'max_execution_time=180',
    '-S', "0.0.0.0:$Port",
    '-t', 'public',
    $router
)

$process = Start-Process -FilePath $php `
    -ArgumentList $arguments `
    -WorkingDirectory $appRoot `
    -RedirectStandardOutput $outLog `
    -RedirectStandardError $errLog `
    -PassThru

$process.Id | Set-Content -Path $pidFile -Encoding ascii

$ready = $false
for ($i = 0; $i -lt 50; $i++) {
    Start-Sleep -Milliseconds 200
    try {
        $response = & curl.exe -sS --max-time 5 -o NUL -w '%{http_code}' "http://127.0.0.1:$Port/"
        if ($LASTEXITCODE -eq 0 -and $response -eq '200') {
            $ready = $true
            break
        }
    } catch {
    }
}

if (-not $ready) {
    if ($process -and -not $process.HasExited) {
        Stop-Process -Id $process.Id -Force
    }
    Remove-Item $pidFile -Force -ErrorAction SilentlyContinue
    throw "No se pudo iniciar Laravel por LAN en el puerto $Port."
}

if ($TryOpenFirewall) {
    $ruleName = "SOFTWARECONTABILIDAD Laravel LAN $Port"
    try {
        if (-not (Get-NetFirewallRule -DisplayName $ruleName -ErrorAction SilentlyContinue)) {
            New-NetFirewallRule -DisplayName $ruleName -Direction Inbound -Action Allow -Protocol TCP -LocalPort $Port -ErrorAction Stop | Out-Null
        }
    } catch {
        Write-Warning "No se pudo crear la regla de firewall automaticamente: $($_.Exception.Message)"
    }
}

$lanIps = Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
    Where-Object {
        $_.IPAddress -and
        $_.IPAddress -notlike '127.*' -and
        $_.IPAddress -notlike '169.254.*'
    } |
    Sort-Object InterfaceMetric, IPAddress |
    Select-Object -ExpandProperty IPAddress -Unique

Write-Host "Servidor Laravel LAN iniciado." -ForegroundColor Green
Write-Host "PID: $($process.Id)"
Write-Host "Puerto: $Port"
Write-Host "URL local: http://127.0.0.1:$Port/"

foreach ($ip in $lanIps) {
    Write-Host "URL LAN: http://$ip`:$Port/"
}

Write-Host "Logs: $outLog"
Write-Host "Errores: $errLog"
Write-Host "Para detenerlo: powershell -ExecutionPolicy Bypass -File .\\scripts\\dev\\stop_laravel_app_lan.ps1 -Port $Port"
