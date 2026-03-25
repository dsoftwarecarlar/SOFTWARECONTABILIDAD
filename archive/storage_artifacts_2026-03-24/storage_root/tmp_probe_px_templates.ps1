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
  $wb = $excel.Workbooks.Open($p, $false, $true)
  try {
    $px = $wb.Worksheets.Item('PX')
    Write-Host "FILE=$p"
    Write-Host (' D3 value=' + [string]$px.Range('D3').Value2 + ' text=' + [string]$px.Range('D3').Text)
    Write-Host (' H3 value=' + [string]$px.Range('H3').Value2 + ' text=' + [string]$px.Range('H3').Text)
    Write-Host (' F49 value=' + [string]$px.Range('F49').Value2 + ' text=' + [string]$px.Range('F49').Text + ' formula=' + [string]$px.Range('F49').Formula)
  } finally { $wb.Close($false) }
}
$excel.Quit()
