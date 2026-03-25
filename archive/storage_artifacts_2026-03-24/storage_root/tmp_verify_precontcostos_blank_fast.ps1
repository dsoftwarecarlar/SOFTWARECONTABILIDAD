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
  $ws = $wb.Worksheets.Item('PrecontabilizacionCostos')
  $main = $ws.Range('B2:J922').Find('*')
  $pivot = $ws.Range('P1:T25').Find('*')
  Write-Host ("{0} main_blank={1} pivot_blank={2}" -f (Split-Path $path -Leaf), ($null -eq $main), ($null -eq $pivot))
  if($null -ne $main){ [void][Runtime.Interopservices.Marshal]::ReleaseComObject($main) }
  if($null -ne $pivot){ [void][Runtime.Interopservices.Marshal]::ReleaseComObject($pivot) }
  $wb.Close($false)
}
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
