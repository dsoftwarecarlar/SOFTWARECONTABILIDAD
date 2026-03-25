param(
    [string]$ListenHost = '127.0.0.1',
    [int]$Port = 8017
)

$ErrorActionPreference = 'Stop'

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..\..')
$php = Join-Path $repoRoot '.tools\php82\php.exe'
$appRoot = Join-Path $repoRoot 'laravel_app'
$router = Join-Path $appRoot 'server.php'
$logOut = Join-Path $appRoot 'storage\logs\laravel_http_servicios.log'
$logErr = Join-Path $appRoot 'storage\logs\laravel_http_servicios.err.log'
$cookie = Join-Path $appRoot 'storage\logs\laravel_servicios.cookies'
$base = "http://$ListenHost`:$Port"
$fixturesRoot = Join-Path $repoRoot 'resources\cxp\servicios_marcas\fixtures'
$files = @(
    @{ field = 'px_file'; path = (Join-Path $fixturesRoot 'detalle-vtas-xliquidar (2).xlsx'); mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' },
    @{ field = 'repventas_file'; path = (Join-Path $fixturesRoot 'RepFacturacionServContabilidad (3).xls'); mime = 'application/vnd.ms-excel' },
    @{ field = 'factura_tyt_file'; path = (Join-Path $fixturesRoot 'archivosasubirtoy\\SERREP_FACTURAS_NAFTOY.TXT'); mime = 'text/plain' },
    @{ field = 'nota_tyt_file'; path = (Join-Path $fixturesRoot 'archivosasubirtoy\\SERREP_NOTACRED_NAFTOY.TXT'); mime = 'text/plain' },
    @{ field = 'mayor_tyt_file'; path = (Join-Path $fixturesRoot 'archivosasubirtoy\\CON_MAYORGEN2TOY.TXT'); mime = 'text/plain' }
)

function Get-ActiveServiciosJobs {
    $jobsDir = Join-Path $repoRoot 'storage\jobs'
    if (-not (Test-Path $jobsDir)) {
        return @()
    }

    $active = @()
    Get-ChildItem -Path $jobsDir -Filter 'servicios_marcas_*.json' | ForEach-Object {
        try {
            $payload = Get-Content -Path $_.FullName -Raw | ConvertFrom-Json
            if ($payload.status -in @('queued', 'running', 'cancel_requested')) {
                $active += [pscustomobject]@{
                    JobId = [string]$payload.job_id
                    Status = [string]$payload.status
                }
            }
        } catch {
        }
    }

    return $active
}

function Get-LatestServiciosJobId {
    $jobsDir = Join-Path $repoRoot 'storage\jobs'
    if (-not (Test-Path $jobsDir)) {
        return ''
    }

    $job = Get-ChildItem -Path $jobsDir -Filter 'servicios_marcas_*.json' |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1

    if (-not $job) {
        return ''
    }

    try {
        $payload = Get-Content -Path $job.FullName -Raw | ConvertFrom-Json
        return [string]$payload.job_id
    } catch {
        return ''
    }
}

function Get-CsrfToken {
    $page = & curl.exe -s -c $cookie "$base/cxp/modules/servicios-marcas"
    if ([string]::IsNullOrWhiteSpace($page)) {
        throw 'No se pudo cargar el modulo Laravel de servicios.'
    }

    $tokenMatch = [regex]::Match($page, 'name="_token" value="([^"]+)"')
    if (-not $tokenMatch.Success) {
        throw 'No se encontro CSRF token para servicios.'
    }

    return $tokenMatch.Groups[1].Value
}

if (-not (Test-Path $php)) {
    throw "No existe PHP 8.2 local en $php"
}

if (-not (Test-Path $router)) {
    throw "No existe router Laravel en $router"
}

foreach ($file in $files) {
    if (-not (Test-Path $file.path)) {
        throw "No existe fixture requerido: $($file.path)"
    }
}

$activeBefore = Get-ActiveServiciosJobs
if ($activeBefore.Count -gt 0) {
    $labels = ($activeBefore | ForEach-Object { "$($_.JobId):$($_.Status)" }) -join ', '
    throw "Servicios por Marca ya tiene jobs activos: $labels"
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
            $status = (& curl.exe -s -o NUL -w '%{http_code}' "$base/").Trim()
            if ($status -match '^\d{3}$' -and $status -ne '000') {
                $ready = $true
                break
            }
        } catch {
        }
        Start-Sleep -Milliseconds 250
    }

    if (-not $ready) {
        throw 'Laravel app no inicio.'
    }

    $token = Get-CsrfToken
    $headerFile = Join-Path $appRoot 'storage\logs\laravel_servicios_headers.txt'
    if (Test-Path $headerFile) {
        Remove-Item $headerFile -Force
    }

    $formArgs = @('-s', '-D', $headerFile, '-o', 'NUL', '-b', $cookie, '-c', $cookie, '-F', "_token=$token", '-F', 'action=process', '-F', 'brand_key=tyt')
    foreach ($file in $files) {
        $formArgs += @('-F', "$($file.field)=@$($file.path);type=$($file.mime)")
    }
    $formArgs += "$base/cxp/modules/servicios-marcas"
    & curl.exe @formArgs | Out-Null

    $headers = Get-Content -Path $headerFile -Raw
    $locationMatch = [regex]::Match($headers, 'Location:\s*(.+?)\r?\n', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $jobId = ''
    if ($locationMatch.Success) {
        $location = $locationMatch.Groups[1].Value.Trim()
        $jobMatch = [regex]::Match($location, '[?&]job=([^&]+)')
        if ($jobMatch.Success) {
            $jobId = [System.Uri]::UnescapeDataString($jobMatch.Groups[1].Value)
        }
    }

    if ([string]::IsNullOrWhiteSpace($jobId)) {
        $jobId = Get-LatestServiciosJobId
    }

    if ([string]::IsNullOrWhiteSpace($jobId)) {
        throw 'No se pudo resolver el job id despues del POST inicial.'
    }

    $statusUrl = "$base/cxp/modules/servicios-marcas/jobs/$jobId"
    $job = $null

    $deadline = (Get-Date).AddMinutes(15)
    do {
        Start-Sleep -Seconds 5
        try {
            $response = Invoke-WebRequest -UseBasicParsing -Uri $statusUrl -TimeoutSec 20
            if ($response.StatusCode -ne 200) {
                throw "Status poll devolvio HTTP $($response.StatusCode)."
            }
            $job = $response.Content | ConvertFrom-Json
        } catch {
            if ((Get-Date) -ge $deadline) {
                throw
            }
            continue
        }
    } while ($job.status -notin @('complete', 'error', 'cancelled') -and (Get-Date) -lt $deadline)

    if ($job.status -ne 'complete') {
        $message = ''
        if ($null -ne $job.error -and [string]$job.error -ne '') {
            $message = [string]$job.error
        } elseif ($null -ne $job.message -and [string]$job.message -ne '') {
            $message = [string]$job.message
        } else {
            $message = 'El job no termino correctamente.'
        }
        throw "Servicios Laravel termino en estado $($job.status): $message"
    }

    if (-not $job.downloads -or $job.downloads.Count -lt 1) {
        throw 'El job completo no devolvio descargas.'
    }

    $downloadUrl = [string]$job.downloads[0].download_url
    if ([string]::IsNullOrWhiteSpace($downloadUrl)) {
        $downloadName = [string]$job.downloads[0].name
        if ([string]::IsNullOrWhiteSpace($downloadName)) {
            throw 'El job completo no devolvio una URL ni nombre de descarga valido.'
        }
        $downloadUrl = "/downloads/$downloadName"
    }
    if ($downloadUrl.StartsWith('/')) {
        $downloadUrl = "$base$downloadUrl"
    }
    $download = Invoke-WebRequest -UseBasicParsing -Uri $downloadUrl -TimeoutSec 30

    [pscustomobject]@{
        Job = $jobId
        Status = [string]$job.status
        Files = $job.downloads.Count
        FirstLabel = [string]$job.downloads[0].label
        FirstFile = [string]$job.downloads[0].name
        DownloadStatus = $download.StatusCode
        Bytes = $download.Content.Length
    } | Format-Table -AutoSize
} finally {
    if ($proc -and -not $proc.HasExited) {
        Stop-Process -Id $proc.Id -Force
    }
}
