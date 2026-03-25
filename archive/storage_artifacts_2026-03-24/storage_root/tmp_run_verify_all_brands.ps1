$outDir='C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\verify_runs'
if (Test-Path $outDir) {
  Get-ChildItem $outDir -File | Remove-Item -Force -ErrorAction SilentlyContinue
} else {
  New-Item -ItemType Directory -Path $outDir | Out-Null
}
$script='C:\xampp\htdocs\SOFTWARECONTABILIDAD\run_servicios_marcas.ps1'
$templateDir='C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\templates'
$rep='C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\fixtures\RepFacturacionServContabilidad (3).xls'
$px='C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\fixtures\detalle-vtas-xliquidar (2).xlsx'
$factChan='C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\fixtures\archivosasubirchan\SERREP_FACTURAS_NAFCHAN.TXT'
$notaChan='C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\fixtures\archivosasubirchan\SERREP_NOTACRED_NAFCHAN.TXT'
$mayorChan='C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\fixtures\archivosasubirchan\CON_MAYORGEN2CHAN.TXT'
$factPeu='C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\fixtures\archivosasubirpeu\SERREP_FACTURAS_NAFPEU.TXT'
$notaPeu='C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\fixtures\archivosasubirpeu\SERREP_NOTACRED_NAFPEU.TXT'
$mayorPeu='C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\fixtures\archivosasubirpeu\CON_MAYORGEN2PEU.TXT'
$factSzk='C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\fixtures\archivosasubirsuz\SERREP_FACTURAS_NAFSUZAMBYRIO.TXT'
$notaSzk='C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\fixtures\archivosasubirsuz\SERREP_NOTACRED_NAFSUZAMBYRI.TXT'
$mayorSzk='C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\fixtures\archivosasubirsuz\CON_MAYORGEN2SUZ.TXT'
$factTyt='C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\fixtures\archivosasubirtoy\SERREP_FACTURAS_NAFTOY.TXT'
$notaTyt='C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\fixtures\archivosasubirtoy\SERREP_NOTACRED_NAFTOY.TXT'
$mayorTyt='C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\fixtures\archivosasubirtoy\CON_MAYORGEN2TOY.TXT'
$sw=[System.Diagnostics.Stopwatch]::StartNew()
& powershell.exe -Sta -NoProfile -ExecutionPolicy Bypass -File $script -InputPath $rep -OutputDir $outDir -TemplateDir $templateDir -RunStamp verifyall_20260320 -RepVtasPath $rep -PxPath $px -FacturaChanganPath $factChan -NotaChanganPath $notaChan -MayorChanganPath $mayorChan -FacturaPeugPath $factPeu -NotaPeugPath $notaPeu -MayorPeugPath $mayorPeu -FacturaSzkPath $factSzk -NotaSzkPath $notaSzk -MayorSzkPath $mayorSzk -FacturaTytPath $factTyt -NotaTytPath $notaTyt -MayorTytPath $mayorTyt 2>&1 | Tee-Object -FilePath (Join-Path $outDir 'worker_all.log')
$sw.Stop()
Write-Host ('ELAPSED_MS=' + $sw.ElapsedMilliseconds)
Write-Host '---FILES---'
Get-ChildItem $outDir | Select-Object Name,Length,LastWriteTime

