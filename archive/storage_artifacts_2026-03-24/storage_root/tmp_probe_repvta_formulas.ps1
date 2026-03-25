$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open('C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\templates\11. Concili. Servicios TYT 2026.xls')
$ws = $wb.Worksheets.Item('REP VTAS')
for ($r=1; $r -le 20; $r++) {
  $vals=@(); for($c=1;$c -le 20;$c++){ $f=$ws.Cells.Item($r,$c).Formula; if($f -is [string] -and $f -match 'Precontabilizacion|COSTO|ESTADISTICAS'){ $vals += ("$c=$f") } }
  if($vals.Count -gt 0){ Write-Host ("R$r " + ($vals -join ' | ')) }
}
$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
