$excel = New-Object -ComObject Excel.Application
$excel.Visible=$false
$excel.DisplayAlerts=$false
try {
  $wb = $excel.Workbooks.Open('C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\templates\11. Concili. Servicios TYT 2026.xls',$false,$true)
  foreach ($ws in @($wb.Worksheets)) {
    $used = $ws.UsedRange
    $lastRow = [int]($used.Row + $used.Rows.Count - 1)
    $lastCol = [int]($used.Column + $used.Columns.Count - 1)
    for ($r=1; $r -le $lastRow; $r++) {
      for ($c=1; $c -le $lastCol; $c++) {
        $cell = $ws.Cells.Item($r,$c)
        $text = [string]$cell.Text
        $formula = [string]$cell.Formula
        if ($text -match '040101110002|SIN IVA' -or $formula -match '040101110002|SIN IVA') {
          Write-Output ('{0}!{1}: TEXT=[{2}] FORMULA=[{3}]' -f $ws.Name,$cell.Address($false,$false),$text,$formula)
        }
        [void][Runtime.Interopservices.Marshal]::ReleaseComObject($cell)
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
