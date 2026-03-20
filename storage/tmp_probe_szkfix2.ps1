$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open('C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_szk_szkfix2_20260319.xls',$false,$true)
try {
  $rep = $wb.Worksheets.Item('REP FACTURACIÓN')
  $nc = $wb.Worksheets.Item('NOTA DE CREDITO')
  $rv = $wb.Worksheets.Item('REP VTAS')
  $pre = $wb.Worksheets.Item('PrecontabilizacionVentas')
  Write-Host ('V7=' + [string]$pre.Range('V7').Text + ' | ' + [string]$pre.Range('V7').Formula)
  Write-Host ('V10=' + [string]$pre.Range('V10').Text + ' | ' + [string]$pre.Range('V10').Formula)
  foreach ($addr in @('D9','E9')) { $c=$rep.Range($addr); Write-Host ('REP ' + $addr + '=' + [string]$c.Text + ' | ' + [string]$c.Formula) }
  foreach ($addr in @('F4','G4')) { $c=$nc.Range($addr); Write-Host ('NC ' + $addr + '=' + [string]$c.Text + ' | ' + [string]$c.Formula) }
  foreach ($addr in @('N2','N3','N4','N5')) { $c=$rv.Range($addr); Write-Host ('RV ' + $addr + '=' + [string]$c.Text + ' | ' + [string]$c.Formula) }
} finally { $wb.Close($false); $excel.Quit() }
