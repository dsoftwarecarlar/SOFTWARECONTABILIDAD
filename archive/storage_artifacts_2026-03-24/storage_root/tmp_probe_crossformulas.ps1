$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$files = @(
 'C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_changan_mirrorallfix_20260319.xls',
 'C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_peug_mirrorallfix_20260319.xls',
 'C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_szk_mirrorall_20260319.xls',
 'C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_tyt_mirrorall_20260319.xls'
)
foreach ($p in $files) {
  $wb = $excel.Workbooks.Open($p,$false,$true)
  try {
    Write-Host "FILE=$p"
    $rep = $wb.Worksheets.Item('REP FACTURACIÓN')
    $nc = $wb.Worksheets.Item('NOTA DE CREDITO')
    $rv = $wb.Worksheets.Item('REP VTAS')
    foreach ($addr in @('D9','E9')) { $c=$rep.Range($addr); Write-Host (' REP!' + $addr + ' text=' + [string]$c.Text + ' formula=' + [string]$c.Formula) }
    foreach ($addr in @('F4','G4')) { $c=$nc.Range($addr); Write-Host (' NC!' + $addr + ' text=' + [string]$c.Text + ' formula=' + [string]$c.Formula) }
    foreach ($addr in @('N2','N3','N4','N5')) { $c=$rv.Range($addr); Write-Host (' RV!' + $addr + ' text=' + [string]$c.Text + ' formula=' + [string]$c.Formula) }
  } finally { $wb.Close($false) }
}
$excel.Quit()
