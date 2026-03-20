$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$paths = @(
  'C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\templates\11. Concili. Servicios CHANGAN  2026.xls',
  'C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\templates\11. Concili. Servicios PEUG  2026.xls',
  'C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\templates\11. Concili. Servicios SZK  2026.xls',
  'C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\templates\11. Concili. Servicios TYT 2026.xls'
)
foreach ($path in $paths) {
  $wb = $excel.Workbooks.Open($path)
  Write-Host "=== $(Split-Path $path -Leaf) ==="
  foreach ($name in 'PrecontabilizacionCostos','PrecontabilizacionCostos (2)') {
    $ws = $wb.Worksheets.Item($name)
    $used = $ws.UsedRange
    Write-Host "[$name] UsedRange Rows=$($used.Rows.Count) Cols=$($used.Columns.Count)"
    for ($r = 1; $r -le [Math]::Min($used.Rows.Count, 25); $r++) {
      $vals = @()
      for ($c = 1; $c -le [Math]::Min($used.Columns.Count, 12); $c++) {
        $v = $ws.Cells.Item($r,$c).Text
        if ($v -ne '') { $vals += ("$c=$v") }
      }
      if ($vals.Count -gt 0) { Write-Host ("R$r " + ($vals -join ' | ')) }
    }
  }
  $wb.Close($false)
}
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
