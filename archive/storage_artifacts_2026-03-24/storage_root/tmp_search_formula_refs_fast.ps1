$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open('C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\templates\11. Concili. Servicios TYT 2026.xls',$false,$true)
try {
  foreach ($sheetName in @('REP FACTURACIÓN','NOTA DE CREDITO','PX','PrecontabilizacionVentas','REP VTAS','MAY VTAS','COSTO','PrecontabilizacionCostos (2)','PrecontabilizacionCostos','ESTADISTICAS')) {
    $ws = $wb.Worksheets.Item($sheetName)
    Write-Host ('SHEET=' + $sheetName)
    try {
      $formulaCells = $ws.UsedRange.SpecialCells(-4123)
      foreach ($cell in $formulaCells) {
        $f = [string]$cell.Formula
        if ($f -like '*PrecontabilizacionCostos*' -or $f -like '*PrecontabilizacionVentas*' -or $f -like '*PrecontabilizacionCostos (2)*') {
          Write-Host (' ' + $cell.Address($false,$false) + ' -> ' + $f)
        }
      }
    } catch {}
  }
} finally { $wb.Close($false); $excel.Quit() }
