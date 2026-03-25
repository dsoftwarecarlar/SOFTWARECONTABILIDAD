$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$files = @(
  @{Brand='CHANGAN'; Path='C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_changan_mirrorall_20260319.xls'},
  @{Brand='PEUGEOT'; Path='C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_peug_mirrorall_20260319.xls'},
  @{Brand='SUZUKI'; Path='C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_szk_mirrorall_20260319.xls'},
  @{Brand='MATRIZ'; Path='C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_tyt_mirrorall_20260319.xls'}
)
foreach ($f in $files) {
  $wb = $excel.Workbooks.Open($f.Path,$false,$true)
  try {
    $ws = $wb.Worksheets.Item('PrecontabilizacionVentas')
    Write-Host ('BRAND=' + $f.Brand)
    for ($r=2; $r -le 15; $r++) {
      $acc = [string]$ws.Cells.Item($r,5).Text
      $desc = [string]$ws.Cells.Item($r,6).Text
      $deb = [string]$ws.Cells.Item($r,8).Text
      $cred = [string]$ws.Cells.Item($r,9).Text
      if ($acc -ne '' -or $desc -ne '' -or $deb -ne '' -or $cred -ne '') {
        Write-Host ($r.ToString() + ': ' + $acc + ' | ' + $desc + ' | D=' + $deb + ' | C=' + $cred)
      }
    }
  } finally { $wb.Close($false) }
}
$excel.Quit()
