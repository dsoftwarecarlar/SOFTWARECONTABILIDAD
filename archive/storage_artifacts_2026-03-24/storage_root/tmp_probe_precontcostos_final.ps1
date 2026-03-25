$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$files = @(
 'C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_changan_finalaudit2_changan_20260319.xls',
 'C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_peug_finalaudit2_peug_20260319.xls',
 'C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_szk_finalaudit2_szk_20260319.xls',
 'C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_tyt_finalaudit_tyt_20260319.xls'
)
foreach($p in $files){
 $wb=$excel.Workbooks.Open($p,$false,$true)
 try{
  $ws=$wb.Worksheets.Item('PrecontabilizacionCostos')
  Write-Host ('FILE=' + $p)
  for($r=1;$r -le 8;$r++){
   $vals=@(); for($c=1;$c -le 10;$c++){$vals += [string]$ws.Cells.Item($r,$c).Text}
   Write-Host ($r.ToString()+': '+($vals -join ' | '))
  }
 } finally { $wb.Close($false) }
}
$excel.Quit()
