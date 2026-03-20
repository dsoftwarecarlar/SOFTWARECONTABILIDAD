$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open('C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_szk_mirrorall_20260319.xls', $false, $true)
try {
  $ws = $wb.Worksheets.Item('REP FACTURACIÓN')
  $cell = $ws.Range('G26')
  Write-Host ('Value2=' + [string]$cell.Value2)
  Write-Host ('Text=' + [string]$cell.Text)
  Write-Host ('Formula=' + [string]$cell.Formula)
  Write-Host ('NumberFormat=' + [string]$cell.NumberFormat)
  Write-Host ('HasFormula=' + [string]$cell.HasFormula)
} finally { $wb.Close($false); $excel.Quit() }
