$excel = New-Object -ComObject Excel.Application
$excel.Visible=$false
$excel.DisplayAlerts=$false
try {
  $wb = $excel.Workbooks.Open('C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\templates\11. Concili. Servicios TYT 2026.xls',$false,$true)
  foreach ($ws in @($wb.Worksheets)) {
    $used = $ws.UsedRange
    foreach ($needle in @('040101110002','SIN IVA')) {
      $first = $used.Find($needle)
      if ($null -ne $first) {
        $addr = $first.Address($false,$false)
        $current = $first
        do {
          Write-Output ('{0}!{1}: TEXT=[{2}] FORMULA=[{3}]' -f $ws.Name,$current.Address($false,$false),$current.Text,$current.Formula)
          $current = $used.FindNext($current)
        } while ($null -ne $current -and $current.Address($false,$false) -ne $addr)
      }
    }
    [void][Runtime.Interopservices.Marshal]::ReleaseComObject($used)
    [void][Runtime.Interopservices.Marshal]::ReleaseComObject($ws)
  }
  $wb.Close($false)
  [void][Runtime.Interopservices.Marshal]::ReleaseComObject($wb)
}
finally {
  $excel.Quit()
  [void][Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
  [gc]::Collect(); [gc]::WaitForPendingFinalizers()
}
