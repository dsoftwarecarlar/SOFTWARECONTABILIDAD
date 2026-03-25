$outDir='C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\perf_runs'
if (Test-Path $outDir) {
  Get-ChildItem $outDir -File | Remove-Item -Force -ErrorAction SilentlyContinue
} else {
  New-Item -ItemType Directory -Path $outDir | Out-Null
}
$script='C:\xampp\htdocs\SOFTWARECONTABILIDAD\run_servicios_marcas.ps1'
$templateDir='C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\templates'
$rep='C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\fixtures\RepFacturacionServContabilidad (3).xls'
$px='C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\fixtures\detalle-vtas-xliquidar (2).xlsx'
$fact='C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\fixtures\archivosasubirtoy\SERREP_FACTURAS_NAFTOY.TXT'
$nota='C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\fixtures\archivosasubirtoy\SERREP_NOTACRED_NAFTOY.TXT'
$mayor='C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\fixtures\archivosasubirtoy\CON_MAYORGEN2TOY.TXT'
$sw=[System.Diagnostics.Stopwatch]::StartNew()
& powershell.exe -Sta -NoProfile -ExecutionPolicy Bypass -File $script -InputPath $rep -OutputDir $outDir -TemplateDir $templateDir -RunStamp perfworker_tyt_20260320 -BrandKey tyt -RepVtasPath $rep -PxPath $px -FacturaTytPath $fact -NotaTytPath $nota -MayorTytPath $mayor 2>&1 | Tee-Object -FilePath (Join-Path $outDir 'worker_tyt.log')
$sw.Stop()
Write-Host ('ELAPSED_MS=' + $sw.ElapsedMilliseconds)
Write-Host '---FILES---'
Get-ChildItem $outDir | Select-Object Name,Length,LastWriteTime
