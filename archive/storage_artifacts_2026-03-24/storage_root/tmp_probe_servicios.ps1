$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$paths = @(
  'C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_changan_mirrorall_20260319.xls',
  'C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_peug_mirrorall_20260319.xls',
  'C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_szk_mirrorall_20260319.xls',
  'C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_tyt_mirrorall_20260319.xls'
)
foreach ($p in $paths) {
  $wb = $excel.Workbooks.Open($p, $false, $true)
  try {
    $px = $wb.Worksheets.Item('PX')
    $pre = $wb.Worksheets.Item('PrecontabilizacionVentas')
    Write-Host "FILE=$p"
    Write-Host (' PX!D5 value=' + [string]$px.Range('D5').Value2 + ' text=' + [string]$px.Range('D5').Text + ' formula=' + [string]$px.Range('D5').Formula)
    Write-Host (' PX!H3 value=' + [string]$px.Range('H3').Value2 + ' text=' + [string]$px.Range('H3').Text + ' formula=' + [string]$px.Range('H3').Formula)
    Write-Host (' PRE!V7 value=' + [string]$pre.Range('V7').Value2 + ' text=' + [string]$pre.Range('V7').Text + ' formula=' + [string]$pre.Range('V7').Formula)
    Write-Host (' PRE!V10 value=' + [string]$pre.Range('V10').Value2 + ' text=' + [string]$pre.Range('V10').Text + ' formula=' + [string]$pre.Range('V10').Formula)
  } finally { $wb.Close($false) }
}
$excel.Quit()
