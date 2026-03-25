$ErrorActionPreference = 'Stop'
$repoRoot = 'C:\xampp\htdocs\SOFTWARECONTABILIDAD'
$php = Join-Path $repoRoot '.tools\php82\php.exe'
$appRoot = Join-Path $repoRoot 'laravel_app'
$publicDir = Join-Path $appRoot 'public'
$listenHost = '127.0.0.1'
$port = 8014
$base = "http://$listenHost`:$port"
$logOut = Join-Path $appRoot 'storage\logs\laravel_http_window1.log'
$logErr = Join-Path $appRoot 'storage\logs\laravel_http_window1.err.log'
$cookie = Join-Path $appRoot 'storage\logs\laravel_window1.cookies'
$fixtures = @{
    accion1 = @(Join-Path $repoRoot 'resources\cxp\acciones\fixtures\CXPREP_docproveedor.pdf')
    accion2 = @(Join-Path $repoRoot 'resources\cxp\acciones\fixtures\CXPREP_RET_GENERALACCION2.txt')
    accion3 = @(Join-Path $repoRoot 'resources\cxp\acciones\fixtures\CON_MAYORGEN2ACCION3.txt')
    accion4 = @(Join-Path $repoRoot 'resources\cxp\acciones\fixtures\CON_MAYORGEN2IVAACCION4.TXT')
    'consolidado-acciones' = @()
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

$proc = Start-Process -FilePath $php -ArgumentList @('-S', "$listenHost`:$port", '-t', $publicDir) -WorkingDirectory $appRoot -RedirectStandardOutput $logOut -RedirectStandardError $logErr -PassThru
try {
    if (Test-Path $cookie) { Remove-Item $cookie -Force }
    $ready = $false
    for ($i = 0; $i -lt 60; $i++) {
        try {
            $resp = Invoke-WebRequest -UseBasicParsing -Uri "$base/" -TimeoutSec 2
            if ($resp.StatusCode -eq 200) { $ready = $true; break }
        } catch {}
        Start-Sleep -Milliseconds 200
    }
    if (-not $ready) { throw 'Laravel app no inicio.' }

    $results = foreach ($slug in @('accion1','accion2','accion3','accion4','consolidado-acciones')) {
        $token = Get-CsrfToken $slug
        $formArgs = @('-s', '-L', '-b', $cookie, '-c', $cookie, '-F', "_token=$token")
        foreach ($fixture in $fixtures[$slug]) {
            $field = if ($slug -eq 'accion3') { 'source_files[]' } else { 'source_files' }
            $mime = if ($fixture -like '*.pdf') { 'application/pdf' } else { 'text/plain' }
            $formArgs += @('-F', "$field=@$fixture;type=$mime")
        }
        $formArgs += "$base/cxp/modules/$slug"
        $response = & curl.exe @formArgs
        if ([string]::IsNullOrWhiteSpace($response)) {
            throw "POST $slug no devolvio contenido."
        }

        $nameMatch = [regex]::Match($response, '<code>([^<]+\.xlsx)</code>')
        [pscustomobject]@{
            Slug = $slug
            Success = [bool]($response -match 'Salida generada')
            Excel = $(if ($nameMatch.Success) { $nameMatch.Groups[1].Value } else { $null })
        }
    }

    $results | ConvertTo-Json -Depth 3 | Write-Output
} finally {
    if ($proc -and -not $proc.HasExited) {
        Stop-Process -Id $proc.Id -Force
    }
}
