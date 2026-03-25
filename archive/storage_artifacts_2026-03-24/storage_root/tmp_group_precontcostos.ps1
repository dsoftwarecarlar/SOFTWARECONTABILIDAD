$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open('C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\templates\11. Concili. Servicios TYT 2026.xls')
$ws = $wb.Worksheets.Item('PrecontabilizacionCostos')
$used = $ws.UsedRange
$lastRow = $used.Row + $used.Rows.Count - 1
$groups = @{}
for ($r=2; $r -le $lastRow; $r++) {
  $ag = ($ws.Cells.Item($r,2).Text).Trim()
  $line = ($ws.Cells.Item($r,3).Text).Trim()
  $num = ($ws.Cells.Item($r,5).Text).Trim()
  $acct = ($ws.Cells.Item($r,4).Text).Trim()
  if ($acct -eq '') { continue }
  $key = "$ag|$line"
  if (-not $groups.ContainsKey($key)) {
    $groups[$key] = [ordered]@{First=$r; Last=$r; Count=0; Numbers=@{}; Accounts=@{}}
  }
  $g = $groups[$key]
  $g.Last = $r
  $g.Count++
  if (-not $g.Numbers.ContainsKey($num)) { $g.Numbers[$num] = 0 }
  $g.Numbers[$num]++
  if (-not $g.Accounts.ContainsKey($acct)) { $g.Accounts[$acct] = 0 }
  $g.Accounts[$acct]++
}
foreach ($entry in $groups.GetEnumerator() | Sort-Object Name) {
  $nums = ($entry.Value.Numbers.Keys | Sort-Object {[double]($_ -replace ',','.')}) -join ','
  $accounts = ($entry.Value.Accounts.Keys | Sort-Object | Select-Object -First 12) -join ','
  Write-Host ("{0} rows={1} range={2}-{3} nums=[{4}] accounts=[{5}]" -f $entry.Key, $entry.Value.Count, $entry.Value.First, $entry.Value.Last, $nums, $accounts)
}
$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
