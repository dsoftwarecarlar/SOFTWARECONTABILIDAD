$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$files = @(
  @{Brand='CHANGAN'; Path='C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_changan_finalaudit2_changan_20260319.xls'; Main='VENTAS'},
  @{Brand='PEUGEOT'; Path='C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_peug_finalaudit2_peug_20260319.xls'; Main='VENTAS'},
  @{Brand='SUZUKI'; Path='C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_szk_finalaudit2_szk_20260319.xls'; Main='VENTAS'},
  @{Brand='MATRIZ'; Path='C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_tyt_finalaudit_tyt_20260319.xls'; Main='MAY VTAS'}
)
foreach ($item in $files) {
  $wb = $excel.Workbooks.Open($item.Path,$false,$true)
  try {
    $rep = $wb.Worksheets.Item('REP FACTURACIÓN')
    $nc = $wb.Worksheets.Item('NOTA DE CREDITO')
    $rv = $wb.Worksheets.Item('REP VTAS')
    $pre = $wb.Worksheets.Item('PrecontabilizacionVentas')
    Write-Host ('BRAND=' + $item.Brand)
    foreach ($addr in @('D9','E9')) { $c=$rep.Range($addr); Write-Host (' REP!' + $addr + '=' + [string]$c.Text) }
    foreach ($addr in @('F4','G4')) { $c=$nc.Range($addr); Write-Host (' NC!' + $addr + '=' + [string]$c.Text) }
    foreach ($addr in @('V7','V10')) { $c=$pre.Range($addr); Write-Host (' PRE!' + $addr + '=' + [string]$c.Text + ' | ' + [string]$c.Formula) }
    foreach ($addr in @('N2','N3','N4','N5')) { $c=$rv.Range($addr); Write-Host (' RV!' + $addr + '=' + [string]$c.Text) }
  } finally { $wb.Close($false) }
}
$excel.Quit()
