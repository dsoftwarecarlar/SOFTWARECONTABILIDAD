$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open('C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_tyt_finalaudit_tyt_20260319.xls')
$ws = $wb.Worksheets.Item('PrecontabilizacionCostos')
$used = $ws.UsedRange
Write-Host "UsedRange Rows=$($used.Rows.Count) Cols=$($used.Columns.Count)"
for ($r=1; $r -le 25; $r++) {
  $vals = @()
  foreach ($c in 2..10) { $text = $ws.Cells.Item($r,$c).Text; if ($text -ne '') { $vals += ("$c=$text") } }
  if ($vals.Count -gt 0) { Write-Host ("R$r " + ($vals -join ' | ')) }
}
$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
