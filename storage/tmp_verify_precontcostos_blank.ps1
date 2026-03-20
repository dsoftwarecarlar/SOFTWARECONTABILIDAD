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
  $hasDataMain = $false
  for($r=2; $r -le 922 -and -not $hasDataMain; $r++){
    for($c=2; $c -le 10; $c++){
      if((($ws.Cells.Item($r,$c).Text).Trim()) -ne ''){ $hasDataMain = $true; break }
    }
  }
  $hasDataPivot = $false
  for($r=1; $r -le 25 -and -not $hasDataPivot; $r++){
    for($c=16; $c -le 20; $c++){
      if((($ws.Cells.Item($r,$c).Text).Trim()) -ne ''){ $hasDataPivot = $true; break }
    }
  }
  Write-Host ("{0} main_blank={1} pivot_blank={2}" -f (Split-Path $path -Leaf), (-not $hasDataMain), (-not $hasDataPivot))
  $wb.Close($false)
}
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
