$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open('C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_tyt_finalaudit_tyt_20260319.xls')
$ws = $wb.Worksheets.Item('REP VTAS')
$lastRow = 350
for ($r=2; $r -le 20; $r++) {
  $vals = @()
  foreach ($c in 1..8) { $text = $ws.Cells.Item($r,$c).Text; if ($text -ne '') { $vals += ("$c=$text") } }
  if ($vals.Count -gt 0) { Write-Host ("R$r " + ($vals -join ' | ')) }
}
$groups=@{}
for ($r=2; $r -le $lastRow; $r++) {
  $center=($ws.Cells.Item($r,2).Text).Trim(" '")
  $advisor=($ws.Cells.Item($r,4).Text).Trim()
  $line=($ws.Cells.Item($r,5).Text).Trim()
  if($center -eq '' -or $center -eq 'CENTRO'){ continue }
  $key="$center|$advisor|$line"
  if(-not $groups.ContainsKey($key)){ $groups[$key]=0 }
  $groups[$key]++
}
Write-Host '--- groups ---'
$groups.GetEnumerator() | Sort-Object Name | ForEach-Object { Write-Host ("{0} count={1}" -f $_.Name,$_.Value) }
$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
