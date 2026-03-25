$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open('C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_changan_mirrorall_20260319.xls', $false, $true)
try {
  $ws = $wb.Worksheets.Item('REP FACTURACIÓN')
  Write-Host ('F17=' + [string]$ws.Range('F17').Text)
  Write-Host ('G17=' + [string]$ws.Range('G17').Text)
  Write-Host ('Q17=' + [string]$ws.Range('Q17').Text)
} finally { $wb.Close($false); $excel.Quit() }
