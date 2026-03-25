param([string]$TemplatePath,[string]$OutputPath,[string]$SheetName,[string]$RangeRef)
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
  $tpl = $excel.Workbooks.Open($TemplatePath, $false, $true)
  $out = $excel.Workbooks.Open($OutputPath, $false, $true)
  $tplSheet = $tpl.Worksheets.Item($SheetName)
  $outSheet = $out.Worksheets.Item($SheetName)
  $tplRange = $tplSheet.Range($RangeRef)
  $outRange = $outSheet.Range($RangeRef)
  $rows = $tplRange.Rows.Count
  $cols = $tplRange.Columns.Count
  for ($r = 1; $r -le $rows; $r++) {
    $parts = @()
    for ($c = 1; $c -le $cols; $c++) {
      $tplCell = $tplRange.Cells.Item($r,$c)
      $outCell = $outRange.Cells.Item($r,$c)
      $addr = $tplCell.Address($false,$false)
      $parts += ('{0}: T=[{1}] F=[{2}] | O=[{3}] F=[{4}]' -f $addr, ($tplCell.Text -replace '\r|\n',' '), ($tplCell.Formula -replace '\r|\n',' '), ($outCell.Text -replace '\r|\n',' '), ($outCell.Formula -replace '\r|\n',' '))
      [void][Runtime.Interopservices.Marshal]::ReleaseComObject($tplCell)
      [void][Runtime.Interopservices.Marshal]::ReleaseComObject($outCell)
    }
    Write-Output ($parts -join ' || ')
  }
  [void][Runtime.Interopservices.Marshal]::ReleaseComObject($tplRange)
  [void][Runtime.Interopservices.Marshal]::ReleaseComObject($outRange)
  [void][Runtime.Interopservices.Marshal]::ReleaseComObject($tplSheet)
  [void][Runtime.Interopservices.Marshal]::ReleaseComObject($outSheet)
  $tpl.Close($false)
  $out.Close($false)
  [void][Runtime.Interopservices.Marshal]::ReleaseComObject($tpl)
  [void][Runtime.Interopservices.Marshal]::ReleaseComObject($out)
}
finally {
  $excel.Quit()
  [void][Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
  [gc]::Collect(); [gc]::WaitForPendingFinalizers()
}
