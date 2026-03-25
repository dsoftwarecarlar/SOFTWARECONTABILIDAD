param(
    [string]$ListenHost = '127.0.0.1',
    [int]$Port = 8027
)

$ErrorActionPreference = 'Stop'

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..\..')
$php = Join-Path $repoRoot '.tools\php82\php.exe'
$appRoot = Join-Path $repoRoot 'laravel_app'
$router = Join-Path $appRoot 'server.php'
$cookie = Join-Path $appRoot 'storage\logs\inspect_bundle.cookies'
$base = "http://$ListenHost`:$Port"
$responseDump = Join-Path $repoRoot 'storage\logs\inspect_bundle_response.html'

$fixtures = @{
    accion1 = @(Join-Path $repoRoot 'resources\cxp\acciones\fixtures\CXPREP_docproveedor.pdf')
    accion2 = @(Join-Path $repoRoot 'resources\cxp\acciones\fixtures\CXPREP_RET_GENERALACCION2.txt')
    accion3 = @(Join-Path $repoRoot 'resources\cxp\acciones\fixtures\CON_MAYORGEN2ACCION3.txt')
    accion4 = @(Join-Path $repoRoot 'resources\cxp\acciones\fixtures\CON_MAYORGEN2IVAACCION4.TXT')
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

function Get-LatestConsolidatedExport {
    $outputsDir = Join-Path $repoRoot 'storage\outputs'
    if (-not (Test-Path $outputsDir)) {
        return $null
    }

    return Get-ChildItem -Path $outputsDir -Filter 'acciones_resumen_*.xlsx' |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
}

$proc = Start-Process -FilePath $php `
    -ArgumentList @('-S', "$ListenHost`:$Port", $router) `
    -WorkingDirectory $appRoot `
    -PassThru

try {
    New-Item -ItemType Directory -Path (Join-Path $repoRoot 'storage\logs') -Force | Out-Null
    if (Test-Path $cookie) {
        Remove-Item $cookie -Force
    }

    for ($i = 0; $i -lt 60; $i++) {
        try {
            $resp = Invoke-WebRequest -UseBasicParsing -Uri "$base/" -TimeoutSec 2
            if ($resp.StatusCode -eq 200) {
                break
            }
        } catch {
        }
        Start-Sleep -Milliseconds 200
    }

    foreach ($slug in @('accion1', 'accion2', 'accion3', 'accion4')) {
        $token = Get-CsrfToken $slug
        $formArgs = @('-s', '-L', '-b', $cookie, '-c', $cookie, '-F', "_token=$token")
        foreach ($fixture in $fixtures[$slug]) {
            $field = if ($slug -eq 'accion3') { 'source_files[]' } else { 'source_files' }
            $mime = if ($fixture -like '*.pdf') { 'application/pdf' } else { 'text/plain' }
            $formArgs += @('-F', "$field=@$fixture;type=$mime")
        }
        $formArgs += "$base/cxp/modules/$slug"
        & curl.exe @formArgs | Out-Null
    }

    $before = Get-LatestConsolidatedExport
    $token = Get-CsrfToken 'consolidado-acciones'
    $response = & curl.exe -s -L -b $cookie -c $cookie -F "_token=$token" "$base/cxp/modules/consolidado-acciones"
    $after = Get-LatestConsolidatedExport

    Set-Content -Path $responseDump -Value $response -Encoding UTF8

    [pscustomobject]@{
        Before = $(if ($before) { $before.Name } else { '' })
        After = $(if ($after) { $after.Name } else { '' })
        Changed = [bool](
            $null -ne $after -and (
                $null -eq $before -or
                [string]$before.FullName -ne [string]$after.FullName -or
                [datetime]$after.LastWriteTimeUtc -gt [datetime]$before.LastWriteTimeUtc
            )
        )
        HasResultText = [bool]($response -match 'Salida generada')
        HasDownload = [bool]($response -match 'href="[^"]*downloads[^"]+"')
        HasErrorText = [bool]($response -match 'No se pudo|Error')
        Dump = $responseDump
    } | Format-List
} finally {
    if ($proc -and -not $proc.HasExited) {
        Stop-Process -Id $proc.Id -Force
    }
}
