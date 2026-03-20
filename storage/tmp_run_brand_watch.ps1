param(
  [string]$BrandKey,
  [string]$RunStamp,
  [string]$ExtraArgs
)
$root = 'C:\xampp\htdocs\SOFTWARECONTABILIDAD'
$outLog = Join-Path $root ("storage\\logs_{0}_{1}.out.log" -f $BrandKey,$RunStamp)
$errLog = Join-Path $root ("storage\\logs_{0}_{1}.err.log" -f $BrandKey,$RunStamp)
if (Test-Path $outLog) { Remove-Item $outLog -Force }
if (Test-Path $errLog) { Remove-Item $errLog -Force }
$cmd = "-ExecutionPolicy Bypass -File `"$root\\scripts\\cxp\\servicios_marcas\\run.ps1`" -InputPath `"$root\\resources\\cxp\\servicios_marcas\\fixtures\\RepFacturacionServContabilidad (3).xls`" -OutputDir `"$root\\storage\\outputs`" -TemplateDir `"$root\\resources\\cxp\\servicios_marcas\\templates`" -RunStamp `"$RunStamp`" -BrandKey `"$BrandKey`" -PxPath `"$root\\resources\\cxp\\servicios_marcas\\fixtures\\detalle-vtas-xliquidar (2).xlsx`" -RepVtasPath `"$root\\resources\\cxp\\servicios_marcas\\fixtures\\RepFacturacionServContabilidad (3).xls`" $ExtraArgs"
$p = Start-Process -FilePath 'C:\WINDOWS\System32\WindowsPowerShell\v1.0\powershell.exe' -ArgumentList $cmd -RedirectStandardOutput $outLog -RedirectStandardError $errLog -PassThru -WorkingDirectory $root
$outputPath = Join-Path $root ("storage\\outputs\\servicios_{0}_{1}.xls" -f $BrandKey,$RunStamp)
$deadline = (Get-Date).AddMinutes(40)
$done = $false
$sawSaveOrValidate = $false
while ((Get-Date) -lt $deadline) {
  Start-Sleep -Seconds 15
  if (Test-Path $outLog) {
    $content = Get-Content $outLog -Raw -ErrorAction SilentlyContinue
    if ($content -match 'INFO\|save_done\|' -or $content -match 'INFO\|validate_done\|') {
      $sawSaveOrValidate = $true
    }
  }
  if ($sawSaveOrValidate -and (Test-Path $outputPath)) {
    $done = $true
    break
  }
  if ($done) { break }
  if (-not (Get-Process -Id $p.Id -ErrorAction SilentlyContinue)) {
    $done = (Test-Path $outputPath)
    break
  }
}
$stillRunning = Get-Process -Id $p.Id -ErrorAction SilentlyContinue
if ($done -and $stillRunning) {
  Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue
  Start-Sleep -Seconds 2
}
Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -eq 0 } | ForEach-Object { Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue }
Write-Host ('OUTPUT=' + $outputPath)
Write-Host ('EXISTS=' + (Test-Path $outputPath))
if (Test-Path $outputPath) { Write-Host ('SIZE=' + (Get-Item $outputPath).Length) }
Write-Host '---OUT---'
if (Test-Path $outLog) { Get-Content $outLog -Tail 40 }
Write-Host '---ERR---'
if (Test-Path $errLog) { Get-Content $errLog -Tail 40 }
