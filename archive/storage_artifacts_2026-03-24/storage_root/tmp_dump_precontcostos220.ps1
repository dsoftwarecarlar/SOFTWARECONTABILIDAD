$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open('C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\templates\11. Concili. Servicios TYT 2026.xls')
$ws = $wb.Worksheets.Item('PrecontabilizacionCostos')
for ($r=1; $r -le 220; $r++) {
  $vals = @()
  foreach ($c in 2..10) { $text = $ws.Cells.Item($r,$c).Text; if ($text -ne '') { $vals += ("$c=$text") } }
  if ($vals.Count -gt 0) { Write-Host ("R$r " + ($vals -join ' | ')) }
}
$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
