$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb=$excel.Workbooks.Open('C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_changan_finalaudit2_changan_20260319.xls',$false,$true)
try{
 $ws=$wb.Worksheets.Item('PrecontabilizacionCostos (2)')
 for($r=1;$r -le 8;$r++){ $vals=@(); for($c=1;$c -le 10;$c++){$vals += [string]$ws.Cells.Item($r,$c).Text}; Write-Host ($r.ToString()+': '+($vals -join ' | ')) }
} finally { $wb.Close($false); $excel.Quit() }
