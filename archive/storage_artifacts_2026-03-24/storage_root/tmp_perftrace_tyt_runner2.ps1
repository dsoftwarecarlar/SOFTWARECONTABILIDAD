$root = 'C:\xampp\htdocs\SOFTWARECONTABILIDAD'
$outLog = Join-Path $root 'storage\perftrace_tyt_20260320.out.log'
$errLog = Join-Path $root 'storage\perftrace_tyt_20260320.err.log'
Remove-Item $outLog,$errLog -Force -ErrorAction SilentlyContinue
$cmd = "-ExecutionPolicy Bypass -File `"$root\scripts\cxp\servicios_marcas\run.ps1`" -InputPath `"$root\resources\cxp\servicios_marcas\fixtures\RepFacturacionServContabilidad (3).xls`" -OutputDir `"$root\storage\outputs`" -TemplateDir `"$root\resources\cxp\servicios_marcas\templates`" -RunStamp `"perftrace_tyt_20260320`" -BrandKey `"tyt`" -PxPath `"$root\resources\cxp\servicios_marcas\fixtures\detalle-vtas-xliquidar (2).xlsx`" -RepVtasPath `"$root\resources\cxp\servicios_marcas\fixtures\RepFacturacionServContabilidad (3).xls`" -FacturaTytPath `"$root\resources\cxp\servicios_marcas\fixtures\archivosasubirtoy\SERREP_FACTURAS_NAFTOY.TXT`" -NotaTytPath `"$root\resources\cxp\servicios_marcas\fixtures\archivosasubirtoy\SERREP_NOTACRED_NAFTOY.TXT`" -MayorTytPath `"$root\resources\cxp\servicios_marcas\fixtures\archivosasubirtoy\CON_MAYORGEN2TOY.TXT`""
$p = Start-Process -FilePath 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe' -ArgumentList $cmd -RedirectStandardOutput $outLog -RedirectStandardError $errLog -WorkingDirectory $root -PassThru
Write-Host ('PID=' + $p.Id)
Start-Sleep -Seconds 480
$alive = Get-Process -Id $p.Id -ErrorAction SilentlyContinue
Write-Host ('ALIVE=' + [bool]$alive)
Write-Host '---OUT TAIL---'
if (Test-Path $outLog) { Get-Content $outLog -Tail 80 }
Write-Host '---ERR TAIL---'
if (Test-Path $errLog) { Get-Content $errLog -Tail 40 }
