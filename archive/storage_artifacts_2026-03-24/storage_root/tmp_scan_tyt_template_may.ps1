$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
  $wb = $excel.Workbooks.Open('C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\templates\11. Concili. Servicios TYT 2026.xls',$false,$true)
  $ws = $wb.Worksheets.Item('MAY VTAS')
  $used = $ws.UsedRange
  $last = [int]($used.Row + $used.Rows.Count - 1)
  for ($r=1; $r -le $last; $r++) {
    $acct = [string]$ws.Cells.Item($r,1).Text
    if ($acct -match '^04\.01\.01\.11\.000[1-4]$') {
      Write-Output ('{0}: A={1} B={2} E={3} F={4} G={5} H={6} I={7} J={8} K={9}' -f $r,$ws.Cells.Item($r,1).Text,$ws.Cells.Item($r,2).Text,$ws.Cells.Item($r,5).Text,$ws.Cells.Item($r,6).Text,$ws.Cells.Item($r,7).Text,$ws.Cells.Item($r,8).Text,$ws.Cells.Item($r,9).Text,$ws.Cells.Item($r,10).Text,$ws.Cells.Item($r,11).Text)
    }
  }
  $wb.Close($false)
  [void][Runtime.Interopservices.Marshal]::ReleaseComObject($used)
  [void][Runtime.Interopservices.Marshal]::ReleaseComObject($ws)
  [void][Runtime.Interopservices.Marshal]::ReleaseComObject($wb)
}
finally {
  $excel.Quit()
  [void][Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
  [gc]::Collect(); [gc]::WaitForPendingFinalizers()
}
