$ErrorActionPreference = 'Stop'
$repoRoot = 'C:\xampp\htdocs\SOFTWARECONTABILIDAD'
$php = Join-Path $repoRoot '.tools\php82\php.exe'
$appRoot = Join-Path $repoRoot 'laravel_app'
$publicDir = Join-Path $appRoot 'public'
$listenHost = '127.0.0.1'
$port = 8015
$base = "http://$listenHost`:$port"
$cookie = Join-Path $appRoot 'storage\logs\laravel_action1_debug.cookies'
$fixture = Join-Path $repoRoot 'resources\cxp\acciones\fixtures\CXPREP_docproveedor.pdf'
$proc = Start-Process -FilePath $php -ArgumentList @('-S', "$listenHost`:$port", '-t', $publicDir) -WorkingDirectory $appRoot -PassThru
try {
    if (Test-Path $cookie) { Remove-Item $cookie -Force }
    for ($i = 0; $i -lt 60; $i++) {
        try {
            $resp = Invoke-WebRequest -UseBasicParsing -Uri "$base/" -TimeoutSec 2
            if ($resp.StatusCode -eq 200) { break }
        } catch {}
        Start-Sleep -Milliseconds 200
    }
    $page = & curl.exe -s -c $cookie "$base/cxp/modules/accion1"
    $token = [regex]::Match($page, 'name="_token" value="([^"]+)"').Groups[1].Value
    $response = & curl.exe -s -L -b $cookie -c $cookie -F "_token=$token" -F "source_files=@$fixture;type=application/pdf" "$base/cxp/modules/accion1"
    $errorMatch = [regex]::Match($response, '<strong>Error</strong>\s*<p>(.*?)</p>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
    [pscustomobject]@{
        HasSuccess = [bool]($response -match 'Salida generada')
        Error = $(if ($errorMatch.Success) { $errorMatch.Groups[1].Value } else { $null })
    } | ConvertTo-Json -Depth 3 | Write-Output
} finally {
    if ($proc -and -not $proc.HasExited) { Stop-Process -Id $proc.Id -Force }
}
