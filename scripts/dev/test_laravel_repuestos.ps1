param(
    [string]$ListenHost = '127.0.0.1',
    [int]$Port = 8016
)

$ErrorActionPreference = 'Stop'

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..\..')
$php = Join-Path $repoRoot '.tools\php82\php.exe'
$appRoot = Join-Path $repoRoot 'laravel_app'
$router = Join-Path $appRoot 'server.php'
$logOut = Join-Path $appRoot 'storage\logs\laravel_http_repuestos.log'
$logErr = Join-Path $appRoot 'storage\logs\laravel_http_repuestos.err.log'
$cookie = Join-Path $appRoot 'storage\logs\laravel_repuestos.cookies'
$base = "http://$ListenHost`:$Port"
$sourceDir = Join-Path $repoRoot 'resources\cxp\repuestos_tytserv\fixtures'
$ncSourceDir = Join-Path $sourceDir 'archivosnc_rep'
$files = @(
    @{ field = 'excel_tyt'; path = (Join-Path $sourceDir 'RepLibroVentasGeneral.xlsx') },
    @{ field = 'excel_nc_tyt'; path = (Join-Path $ncSourceDir 'RepLibroDevolucionesGeneral.xlsx') },
    @{ field = 'excel_peug'; path = (Join-Path $sourceDir 'RepLibroVentasGeneral (1).xlsx') },
    @{ field = 'excel_nc_peug'; path = (Join-Path $ncSourceDir 'RepLibroDevolucionesGeneral (1).xlsx') },
    @{ field = 'excel_chgn'; path = (Join-Path $sourceDir 'RepLibroVentasGeneral (2).xlsx') },
    @{ field = 'excel_nc_chgn'; path = (Join-Path $ncSourceDir 'RepLibroDevolucionesGeneral (2).xlsx') },
    @{ field = 'excel_szk'; path = (Join-Path $sourceDir 'RepLibroVentasGeneral (3).xlsx') },
    @{ field = 'excel_nc_szk'; path = (Join-Path $ncSourceDir 'RepLibroDevolucionesGeneral (3).xlsx') }
)

function Get-CsrfToken {
    $page = & curl.exe -s -c $cookie "$base/cxp/modules/repuestos-tytserv"
    if ([string]::IsNullOrWhiteSpace($page)) {
        throw 'No se pudo cargar el modulo Laravel de repuestos.'
    }

    $tokenMatch = [regex]::Match($page, 'name="_token" value="([^"]+)"')
    if (-not $tokenMatch.Success) {
        throw 'No se encontro CSRF token para repuestos.'
    }

    return $tokenMatch.Groups[1].Value
}

if (-not (Test-Path $php)) {
    throw "No existe PHP 8.2 local en $php"
}

foreach ($file in $files) {
    if (-not (Test-Path $file.path)) {
        throw "No existe fixture requerido: $($file.path)"
    }
}

$proc = Start-Process -FilePath $php `
    -ArgumentList @('-d', 'max_execution_time=180', '-S', "$ListenHost`:$Port", $router) `
    -WorkingDirectory $appRoot `
    -RedirectStandardOutput $logOut `
    -RedirectStandardError $logErr `
    -PassThru

try {
    if (Test-Path $cookie) {
        Remove-Item $cookie -Force
    }

    $ready = $false
    for ($i = 0; $i -lt 60; $i++) {
        try {
            $resp = Invoke-WebRequest -UseBasicParsing -Uri "$base/" -TimeoutSec 2
            if ($resp.StatusCode -eq 200) {
                $ready = $true
                break
            }
        } catch {
        }
        Start-Sleep -Milliseconds 200
    }

    if (-not $ready) {
        throw 'Laravel app no inicio.'
    }

    $token = Get-CsrfToken
    $formArgs = @('-s', '-L', '-b', $cookie, '-c', $cookie, '-F', "_token=$token")
    foreach ($file in $files) {
        $formArgs += @('-F', "$($file.field)=@$($file.path);type=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    }
    $formArgs += "$base/cxp/modules/repuestos-tytserv"

    $response = & curl.exe @formArgs
    if ([string]::IsNullOrWhiteSpace($response)) {
        throw 'POST repuestos no devolvio contenido.'
    }

    $downloadMatch = [regex]::Match($response, 'href="([^"]*downloads[^"]+)"')

    if (-not $downloadMatch.Success) {
        throw 'No se encontro enlace de descarga del Excel generado.'
    }

    $downloadUrl = $downloadMatch.Groups[1].Value
    if ($downloadUrl.StartsWith('/')) {
        $downloadUrl = "$base$downloadUrl"
    }
    $downloadFileName = [System.IO.Path]::GetFileName(([Uri]$downloadUrl).AbsolutePath)
    $download = Invoke-WebRequest -UseBasicParsing -Uri $downloadUrl -TimeoutSec 20

    $result = [pscustomobject]@{
        Success = [bool]($download.StatusCode -eq 200 -and $downloadMatch.Success)
        Excel = $downloadFileName
        DownloadStatus = $download.StatusCode
        Bytes = $download.Content.Length
    }

    $result | Format-Table -AutoSize

    if (-not $result.Success) {
        throw 'Repuestos Laravel no genero una descarga valida.'
    }
} finally {
    if ($proc -and -not $proc.HasExited) {
        Stop-Process -Id $proc.Id -Force
    }
}
