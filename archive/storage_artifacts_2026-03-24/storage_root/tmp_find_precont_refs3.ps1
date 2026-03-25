$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open('C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\templates\11. Concili. Servicios TYT 2026.xls')
foreach ($ws in $wb.Worksheets) {
  $first = $ws.Cells.Find('PrecontabilizacionCostos')
  if ($null -ne $first) {
    $firstAddr = $first.Address()
    $cell = $first
    do {
      Write-Host ("{0}!{1} formula={2}" -f $ws.Name, $cell.Address(), $cell.Formula)
      $cell = $ws.Cells.FindNext($cell)
    } while ($null -ne $cell -and $cell.Address() -ne $firstAddr)
  }
}
$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
