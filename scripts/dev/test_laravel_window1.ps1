param(
    [string]$ListenHost = '127.0.0.1',
    [int]$Port = 8014
)

$ErrorActionPreference = 'Stop'

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..\..')
$php = Join-Path $repoRoot '.tools\php82\php.exe'
$appRoot = Join-Path $repoRoot 'laravel_app'
$router = Join-Path $appRoot 'server.php'
$logOut = Join-Path $appRoot 'storage\logs\laravel_http_window1.log'
$logErr = Join-Path $appRoot 'storage\logs\laravel_http_window1.err.log'
$cookie = Join-Path $appRoot 'storage\logs\laravel_window1.cookies'
$base = "http://$ListenHost`:$Port"
$fixtures = @{
    accion1 = @(Join-Path $repoRoot 'resources\cxp\acciones\fixtures\CXPREP_docproveedor.pdf')
    accion2 = @(Join-Path $repoRoot 'resources\cxp\acciones\fixtures\CXPREP_RET_GENERALACCION2.txt')
    accion3 = @(Join-Path $repoRoot 'resources\cxp\acciones\fixtures\CON_MAYORGEN2ACCION3.txt')
    accion4 = @(Join-Path $repoRoot 'resources\cxp\acciones\fixtures\CON_MAYORGEN2IVAACCION4.TXT')
    'consolidado-acciones' = @()
}

function Get-LatestConsolidatedExport {
    $outputsDir = Join-Path $repoRoot 'storage\outputs'
    if (-not (Test-Path $outputsDir)) {
        return $null
    }

    return Get-ChildItem -Path $outputsDir -Filter 'acciones_resumen_*.xlsx' |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
}

function Test-ConsolidatedChanged {
    param(
        $BeforeExport,
        $AfterExport
    )

    if ($null -eq $AfterExport) {
        return $false
    }

    if ($null -eq $BeforeExport) {
        return $true
    }

    return (
        [string]$AfterExport.FullName -ne [string]$BeforeExport.FullName -or
        [datetime]$AfterExport.LastWriteTimeUtc -gt [datetime]$BeforeExport.LastWriteTimeUtc
    )
}

function Get-CsrfToken([string]$Slug) {
    $page = & curl.exe -s -c $cookie "$base/cxp/modules/$Slug"
    if ([string]::IsNullOrWhiteSpace($page)) {
        throw "No se pudo cargar $Slug."
    }

    $tokenMatch = [regex]::Match($page, 'name="_token" value="([^"]+)"')
    if (-not $tokenMatch.Success) {
        throw "No se encontro CSRF token para $Slug."
    }

    return $tokenMatch.Groups[1].Value
}

if (-not (Test-Path $php)) {
    throw "No existe PHP 8.2 local en $php"
}

$proc = Start-Process -FilePath $php `
    -ArgumentList @('-S', "$ListenHost`:$Port", $router) `
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

    $results = foreach ($slug in @('accion1', 'accion2', 'accion3', 'accion4', 'consolidado-acciones')) {
        $token = Get-CsrfToken $slug
        $formArgs = @('-s', '-L', '-b', $cookie, '-c', $cookie, '-F', "_token=$token")
        foreach ($fixture in $fixtures[$slug]) {
            $field = if ($slug -eq 'accion3') { 'source_files[]' } else { 'source_files' }
            $mime = if ($fixture -like '*.pdf') { 'application/pdf' } else { 'text/plain' }
            $formArgs += @('-F', "$field=@$fixture;type=$mime")
        }
        $formArgs += "$base/cxp/modules/$slug"

        $beforeConsolidated = $null
        if ($slug -eq 'consolidado-acciones') {
            $beforeConsolidated = Get-LatestConsolidatedExport
        }

        $response = & curl.exe @formArgs
        if ([string]::IsNullOrWhiteSpace($response)) {
            throw "POST $slug no devolvio contenido."
        }

        $nameMatch = [regex]::Match($response, '<code>([^<]+\.xlsx)</code>')
        $hasDownload = [regex]::IsMatch($response, 'href="[^"]*downloads[^"]+"')
        $afterConsolidated = $null
        $success = if ($slug -eq 'consolidado-acciones') {
            $newExportDetected = $false
            for ($retry = 0; $retry -lt 60; $retry++) {
                $afterConsolidated = Get-LatestConsolidatedExport
                $newExportDetected = Test-ConsolidatedChanged -BeforeExport $beforeConsolidated -AfterExport $afterConsolidated
                if ($newExportDetected) {
                    break
                }
                Start-Sleep -Seconds 1
            }
            [bool]($newExportDetected -or (($response -match 'Salida generada') -and $hasDownload))
        } else {
            [bool](($response -match 'Salida generada') -and $nameMatch.Success)
        }
        [pscustomobject]@{
            Slug = $slug
            Success = $success
            Excel = $(if ($nameMatch.Success) { $nameMatch.Groups[1].Value } elseif ($slug -eq 'consolidado-acciones' -and $null -ne $afterConsolidated) { $afterConsolidated.Name } else { $null })
        }
    }

    $results | Format-Table -AutoSize

    $failed = @($results | Where-Object { -not $_.Success })
    if ($failed.Count -gt 0) {
        $labels = ($failed | ForEach-Object { $_.Slug }) -join ', '
        throw "Ventana 1 Laravel tuvo modulos sin exito: $labels"
    }
} finally {
    if ($proc -and -not $proc.HasExited) {
        Stop-Process -Id $proc.Id -Force
    }
}
