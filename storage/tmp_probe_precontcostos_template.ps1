$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$paths = @(
  'C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\templates\11. Concili. Servicios CHANGAN  2026.xls',
  'C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\templates\11. Concili. Servicios PEUG  2026.xls',
  'C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\templates\11. Concili. Servicios SZK  2026.xls',
  'C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\templates\11. Concili. Servicios TYT 2026.xls'
)
foreach ($p in $paths) {
  $wb = $excel.Workbooks.Open($p,$false,$true)
  try {
    $ws = $wb.Worksheets.Item('PrecontabilizacionCostos')
    Write-Host "FILE=$p"
    for ($r=1; $r -le 12; $r++) {
      $vals = @()
      for ($c=1; $c -le 20; $c++) {
        $vals += ([string]$ws.Cells.Item($r,$c).Text)
      }
      Write-Host ($r.ToString().PadLeft(2) + ': ' + ($vals -join ' | '))
    }
    Write-Host '---'
  } finally { $wb.Close($false) }
}
$excel.Quit()
