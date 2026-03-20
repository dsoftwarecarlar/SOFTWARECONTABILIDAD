$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$paths = @(
  'C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\templates\11. Concili. Servicios CHANGAN  2026.xls',
  'C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\templates\11. Concili. Servicios TYT 2026.xls'
)
foreach ($path in $paths) {
  $wb = $excel.Workbooks.Open($path)
  Write-Host "=== $(Split-Path $path -Leaf) ==="
  foreach ($ws in $wb.Worksheets) {
    $used = $ws.UsedRange
    $rows = [Math]::Min($used.Rows.Count, 1200)
    $cols = [Math]::Min($used.Columns.Count, 40)
    for ($r=1; $r -le $rows; $r++) {
      for ($c=1; $c -le $cols; $c++) {
        $f = $ws.Cells.Item($r,$c).Formula
        if ($f -is [string] -and $f -match 'PrecontabilizacionCostos') {
          Write-Host ("{0}!{1}{2} = {3}" -f $ws.Name, ([char](64+[math]::Min($c,26))), $r, $f)
        }
      }
    }
  }
  $wb.Close($false)
}
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
