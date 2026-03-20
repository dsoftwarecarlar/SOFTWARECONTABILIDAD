$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$paths = @(
  'C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_changan_mirrorallfix_20260319.xls',
  'C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_peug_mirrorallfix_20260319.xls'
)
foreach ($p in $paths) {
  $wb = $excel.Workbooks.Open($p, $false, $true)
  try {
    $pre = $wb.Worksheets.Item('PrecontabilizacionVentas')
    $costo = $wb.Worksheets.Item('PrecontabilizacionCostos')
    Write-Host "FILE=$p"
    Write-Host (' V7=' + [string]$pre.Range('V7').Text + ' | formula=' + [string]$pre.Range('V7').Formula)
    Write-Host (' V10=' + [string]$pre.Range('V10').Text + ' | formula=' + [string]$pre.Range('V10').Formula)
    Write-Host (' Row2Height=' + [string]$costo.Rows.Item(2).RowHeight + ' Row100Height=' + [string]$costo.Rows.Item(100).RowHeight)
  } finally { $wb.Close($false) }
}
$excel.Quit()
