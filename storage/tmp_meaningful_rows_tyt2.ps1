function Get-LastMeaningfulCell($ws) {
  $after = $ws.Cells.Item(1,1)
  $lastRowCell = $ws.Cells.Find('*', $after, -4123, 2, 1, 2, $false, $false, $false)
  $lastColCell = $ws.Cells.Find('*', $after, -4123, 2, 2, 2, $false, $false, $false)
  $lastRow = if($lastRowCell){$lastRowCell.Row}else{1}
  $lastCol = if($lastColCell){$lastColCell.Column}else{1}
  if($lastRowCell){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($lastRowCell)}
  if($lastColCell){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($lastColCell)}
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($after)
  return @($lastRow,$lastCol)
}
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open('C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\templates\11. Concili. Servicios TYT 2026.xls', $null, $true)
try {
  foreach($name in 'REP FACTURACIÓN','NOTA DE CREDITO','REP VTAS','PX','PrecontabilizacionVentas','PrecontabilizacionCostos','PrecontabilizacionCostos (2)','ESTADISTICAS','COSTO','MAY VTAS') {
    $ws = $wb.Worksheets.Item($name)
    $ur = $ws.UsedRange
    $lastMeaningful = Get-LastMeaningfulCell $ws
    '{0}|usedLastRow={1}|usedLastCol={2}|meaningfulLastRow={3}|meaningfulLastCol={4}' -f $name, ([int]($ur.Row + $ur.Rows.Count -1)), ([int]($ur.Column + $ur.Columns.Count -1)), $lastMeaningful[0], $lastMeaningful[1]
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ur)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ws)
  }
} finally {
  $wb.Close($false)
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
  $excel.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
[GC]::Collect()
