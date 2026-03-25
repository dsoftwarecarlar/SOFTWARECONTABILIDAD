$excel = New-Object -ComObject Excel.Application
$excel.Visible=$false; $excel.DisplayAlerts=$false
$wb=$excel.Workbooks.Open('C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\templates\11. Concili. Servicios CHANGAN  2026.xls',$false,$true)
try{
 $ws=$wb.Worksheets.Item('PrecontabilizacionCostos')
 for($r=9;$r -le 20;$r++){ $vals=@(); for($c=2;$c -le 10;$c++){$vals += [string]$ws.Cells.Item($r,$c).Text}; Write-Host ($r.ToString()+': '+($vals -join ' | ')) }
} finally { $wb.Close($false); $excel.Quit() }
