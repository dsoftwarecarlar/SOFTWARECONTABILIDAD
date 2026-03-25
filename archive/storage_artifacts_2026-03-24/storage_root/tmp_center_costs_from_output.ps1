$path='C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_tyt_finalaudit_tyt_20260319.xls'
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open($path)
$ws = $wb.Worksheets.Item('REP VTAS')
$used = $ws.UsedRange
$lastRow = $used.Row + $used.Rows.Count - 1
$byCenter = @{}
for ($r=2; $r -le $lastRow; $r++) {
  $center = ($ws.Cells.Item($r,2).Text).Trim(" '")
  if ($center -eq '') { continue }
  if (-not $byCenter.ContainsKey($center)) { $byCenter[$center] = [ordered]@{Rows=0; Costo=0.0; Lub=0.0; Acc=0.0; Rep=0.0; Pint=0.0; SubNc=0.0} }
  $g=$byCenter[$center]
  $g.Rows++
  $g.Costo += [double]($ws.Cells.Item($r,21).Value2 -as [double])
  $g.Lub += [double]($ws.Cells.Item($r,22).Value2 -as [double])
  $g.Acc += [double]($ws.Cells.Item($r,23).Value2 -as [double])
  $g.Rep += [double]($ws.Cells.Item($r,24).Value2 -as [double])
  $g.Pint += [double]($ws.Cells.Item($r,25).Value2 -as [double])
  $g.SubNc += [double]($ws.Cells.Item($r,26).Value2 -as [double])
}
$byCenter.GetEnumerator() | Sort-Object Name | ForEach-Object {
  $v=$_.Value
  Write-Host ("{0} rows={1} costo={2:N2} lub={3:N2} acc={4:N2} rep={5:N2} pint={6:N2} subNc={7:N2}" -f $_.Name,$v.Rows,$v.Costo,$v.Lub,$v.Acc,$v.Rep,$v.Pint,$v.SubNc)
}
$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
