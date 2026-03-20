$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open('C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\templates\11. Concili. Servicios TYT 2026.xls')
foreach ($ws in $wb.Worksheets) {
  $range = $ws.Cells
  $first = $range.Find('PrecontabilizacionCostos', $null, -4123, 1, 1, 1, $false)
  if ($null -ne $first) {
    $addr = $first.Address()
    $current = $first
    do {
      Write-Host ("{0}!{1} formula={2}" -f $ws.Name, $current.Address(), $current.Formula)
      $current = $range.FindNext($current)
    } while ($null -ne $current -and $current.Address() -ne $addr)
  }
}
$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
