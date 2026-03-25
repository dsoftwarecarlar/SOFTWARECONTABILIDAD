$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$files = @(
  'C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_changan_finalaudit2_changan_20260319.xls',
  'C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_peug_finalaudit2_peug_20260319.xls',
  'C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_szk_finalaudit2_szk_20260319.xls',
  'C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_tyt_finalaudit_tyt_20260319.xls'
)
foreach($path in $files){
  $wb = $excel.Workbooks.Open($path)
  $ws = $wb.Worksheets.Item('REP VTAS')
  $used = $ws.UsedRange
  $lastRow = [Math]::Min($used.Row + $used.Rows.Count - 1, 400)
  $centers = New-Object System.Collections.Generic.List[string]
  $lines = New-Object System.Collections.Generic.List[string]
  for($r=15; $r -le $lastRow; $r++){
    $center = ($ws.Cells.Item($r,2).Text).Trim(" '")
    $line = ($ws.Cells.Item($r,5).Text).Trim()
    if($center -ne '' -and -not $centers.Contains($center)){ [void]$centers.Add($center) }
    if($line -ne '' -and -not $lines.Contains($line)){ [void]$lines.Add($line) }
  }
  Write-Host ("{0} centers=[{1}] lines=[{2}]" -f (Split-Path $path -Leaf), (($centers | Sort-Object) -join ','), (($lines | Sort-Object) -join ','))
  $wb.Close($false)
}
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
