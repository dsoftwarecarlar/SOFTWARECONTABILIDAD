$ErrorActionPreference = 'Stop'
$repoRoot = 'C:\xampp\htdocs\SOFTWARECONTABILIDAD'
$php = Join-Path $repoRoot '.tools\php82\php.exe'
$appRoot = Join-Path $repoRoot 'laravel_app'
$publicDir = Join-Path $appRoot 'public'
$listenHost = '127.0.0.1'
$port = 8013
$base = "http://$listenHost`:$port"
$logOut = Join-Path $appRoot 'storage\logs\laravel_http_action1.log'
$logErr = Join-Path $appRoot 'storage\logs\laravel_http_action1.err.log'
$cookie = Join-Path $appRoot 'storage\logs\laravel_action1.cookies'
$fixture = Join-Path $repoRoot 'resources\cxp\acciones\fixtures\CXPREP_docproveedor.pdf'
if (Test-Path $cookie) { Remove-Item $cookie -Force }
$proc = Start-Process -FilePath $php -ArgumentList @('-S', "$listenHost`:$port", '-t', $publicDir) -WorkingDirectory $appRoot -RedirectStandardOutput $logOut -RedirectStandardError $logErr -PassThru
try {
    $ready = $false
    for ($i = 0; $i -lt 60; $i++) {
        try {
            $resp = Invoke-WebRequest -UseBasicParsing -Uri "$base/" -TimeoutSec 2
            if ($resp.StatusCode -eq 200) { $ready = $true; break }
        } catch {}
        Start-Sleep -Milliseconds 200
    }
    if (-not $ready) { throw 'Laravel app no inicio.' }

    $page = & curl.exe -s -c $cookie "$base/cxp/modules/accion1"
    if ([string]::IsNullOrWhiteSpace($page)) { throw 'No se pudo cargar accion1.' }

    $tokenMatch = [regex]::Match($page, 'name="_token" value="([^"]+)"')
    if (-not $tokenMatch.Success) { throw 'No se encontro CSRF token.' }
    $token = $tokenMatch.Groups[1].Value

    $response = & curl.exe -s -L -b $cookie -c $cookie -F "_token=$token" -F "source_files=@$fixture;type=application/pdf" "$base/cxp/modules/accion1"
    if ([string]::IsNullOrWhiteSpace($response)) { throw 'POST accion1 no devolvio contenido.' }

    $hasSuccess = $response -match 'Salida generada'
    $downloadMatch = [regex]::Match($response, 'href="([^"]*downloads[^"]+)"')
    $nameMatch = [regex]::Match($response, '<code>([^<]+\.xlsx)</code>')

    [pscustomobject]@{
        Success = $hasSuccess
        Excel = $(if ($nameMatch.Success) { $nameMatch.Groups[1].Value } else { $null })
        Download = $(if ($downloadMatch.Success) { $downloadMatch.Groups[1].Value } else { $null })
    } | ConvertTo-Json -Depth 3 | Write-Output
} finally {
    if ($proc -and -not $proc.HasExited) {
        Stop-Process -Id $proc.Id -Force
    }
}
