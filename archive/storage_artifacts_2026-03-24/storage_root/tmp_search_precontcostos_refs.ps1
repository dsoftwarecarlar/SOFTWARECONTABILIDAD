$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$paths = @(
  'C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\templates\11. Concili. Servicios CHANGAN  2026.xls',
  'C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\templates\11. Concili. Servicios TYT 2026.xls'
)
foreach ($p in $paths) {
  $wb = $excel.Workbooks.Open($p,$false,$true)
  try {
    Write-Host "FILE=$p"
    foreach ($ws in $wb.Worksheets) {
      $used = $ws.UsedRange
      $found = $false
      foreach ($cell in $used) {
        $f = [string]$cell.Formula
        if ($f -like '*PrecontabilizacionCostos*') {
          if (-not $found) { Write-Host (' SHEET=' + $ws.Name); $found = $true }
          Write-Host ('  ' + $cell.Address($false,$false) + ' -> ' + $f)
        }
      }
    }
  } finally { $wb.Close($false) }
}
$excel.Quit()
