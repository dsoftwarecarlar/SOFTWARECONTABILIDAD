$stdout = 'C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\szkdiag_20260319.out.log'
$stderr = 'C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\szkdiag_20260319.err.log'
if (Test-Path $stdout) { Remove-Item $stdout -Force }
if (Test-Path $stderr) { Remove-Item $stderr -Force }
$p = Start-Process -FilePath 'C:\WINDOWS\System32\WindowsPowerShell\v1.0\powershell.exe' -ArgumentList @(
  '-ExecutionPolicy','Bypass','-File','C:\xampp\htdocs\SOFTWARECONTABILIDAD\scripts\cxp\servicios_marcas\run.ps1',
  '-InputPath','C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\fixtures\RepFacturacionServContabilidad (3).xls',
  '-OutputDir','C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs',
  '-TemplateDir','C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\templates',
  '-RunStamp','szkdiag_20260319',
  '-BrandKey','szk',
  '-PxPath','C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\fixtures\detalle-vtas-xliquidar (2).xlsx',
  '-RepVtasPath','C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\fixtures\RepFacturacionServContabilidad (3).xls',
  '-FacturaSzkPath','C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\fixtures\archivosasubirsuz\SERREP_FACTURAS_NAFSUZAMBYRIO.TXT',
  '-NotaSzkPath','C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\fixtures\archivosasubirsuz\SERREP_NOTACRED_NAFSUZAMBYRI.TXT',
  '-MayorSzkPath','C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\fixtures\archivosasubirsuz\CON_MAYORGEN2SUZ.TXT'
) -RedirectStandardOutput $stdout -RedirectStandardError $stderr -PassThru -WorkingDirectory 'C:\xampp\htdocs\SOFTWARECONTABILIDAD'
Start-Sleep -Seconds 90
$running = Get-Process -Id $p.Id -ErrorAction SilentlyContinue
Write-Host ('PID=' + $p.Id + ' RUNNING=' + [bool]$running)
Write-Host '---STDOUT---'
if (Test-Path $stdout) { Get-Content $stdout -Tail 80 }
Write-Host '---STDERR---'
if (Test-Path $stderr) { Get-Content $stderr -Tail 80 }
