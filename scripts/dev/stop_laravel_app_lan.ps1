param(
    [int]$Port = 8085
)

$ErrorActionPreference = 'Stop'

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..\..')
$pidFile = Join-Path $repoRoot "laravel_app\storage\logs\laravel_lan_$Port.pid"

if (-not (Test-Path $pidFile)) {
    Write-Host "No existe PID activo registrado para el puerto $Port."
    exit 0
}

$pid = Get-Content $pidFile -ErrorAction SilentlyContinue | Select-Object -First 1
if ($pid) {
    $process = Get-Process -Id $pid -ErrorAction SilentlyContinue
    if ($process) {
        Stop-Process -Id $pid -Force
        Write-Host "Servidor Laravel LAN detenido (PID $pid)."
    } else {
        Write-Host "El PID $pid ya no estaba activo."
    }
}

Remove-Item $pidFile -Force -ErrorAction SilentlyContinue
