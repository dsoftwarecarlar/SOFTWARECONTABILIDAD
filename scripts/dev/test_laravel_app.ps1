param(
    [string]$ListenHost = '127.0.0.1',
    [int]$Port = 8011
)

$ErrorActionPreference = 'Stop'

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..\..')
$php = Join-Path $repoRoot '.tools\php82\php.exe'
$appRoot = Join-Path $repoRoot 'laravel_app'
$router = Join-Path $appRoot 'server.php'
$logOut = Join-Path $appRoot 'storage\logs\laravel_http_probe.log'
$logErr = Join-Path $appRoot 'storage\logs\laravel_http_probe.err.log'
$urls = @(
    "http://$ListenHost`:$Port/",
    "http://$ListenHost`:$Port/cxp",
    "http://$ListenHost`:$Port/cxp/windows/libro-compras-aclt",
    "http://$ListenHost`:$Port/cxp/windows/conciliacion-servicios-marcas",
    "http://$ListenHost`:$Port/cxp/windows/facturacion-repuestos-tytserv",
    "http://$ListenHost`:$Port/cxp/modules/accion1",
    "http://$ListenHost`:$Port/cxp/modules/accion2",
    "http://$ListenHost`:$Port/cxp/modules/accion3",
    "http://$ListenHost`:$Port/cxp/modules/accion4",
    "http://$ListenHost`:$Port/cxp/modules/consolidado-acciones",
    "http://$ListenHost`:$Port/cxp/modules/servicios-marcas",
    "http://$ListenHost`:$Port/cxp/modules/repuestos-tytserv"
)

function Invoke-HttpText {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Url,

        [int]$TimeoutSec = 5
    )

    $bodyFile = [System.IO.Path]::GetTempFileName()
    try {
        $statusCode = & curl.exe -L -sS --max-time $TimeoutSec -A 'LaravelAppProbe/1.0' -o $bodyFile -w '%{http_code}' $Url
        $curlExitCode = $LASTEXITCODE
        if ($curlExitCode -ne 0 -or -not $statusCode) {
            throw "No se pudo consultar $Url con curl (exit $curlExitCode)."
        }

        [pscustomobject]@{
            StatusCode = [int]$statusCode
            Content = [System.IO.File]::ReadAllText($bodyFile)
        }
    } finally {
        if (Test-Path $bodyFile) {
            Remove-Item $bodyFile -Force -ErrorAction SilentlyContinue
        }
    }
}

if (-not (Test-Path $php)) {
    throw "No existe PHP 8.2 local en $php"
}

$proc = Start-Process -FilePath $php `
    -ArgumentList @('-S', "$ListenHost`:$Port", '-t', 'public', $router) `
    -WorkingDirectory $appRoot `
    -RedirectStandardOutput $logOut `
    -RedirectStandardError $logErr `
    -PassThru

try {
    $ready = $false
    for ($i = 0; $i -lt 50; $i++) {
        try {
            $response = Invoke-HttpText -Url $urls[0] -TimeoutSec 10
            if ($response.StatusCode -eq 200) {
                $ready = $true
                break
            }
        } catch {
        }
        Start-Sleep -Milliseconds 200
    }

    if (-not $ready) {
        throw 'Laravel HTTP probe no pudo iniciar.'
    }

    $results = foreach ($url in $urls) {
        $response = Invoke-HttpText -Url $url -TimeoutSec 15
        if ($response.StatusCode -ne 200) {
            throw "Laravel HTTP probe recibio estado $($response.StatusCode) para $url"
        }
        [pscustomobject]@{
            Url = $url
            Status = $response.StatusCode
            Bytes = $response.Content.Length
        }
    }

    $results | Format-Table -AutoSize
} finally {
    if ($proc -and -not $proc.HasExited) {
        Stop-Process -Id $proc.Id -Force
    }
}
