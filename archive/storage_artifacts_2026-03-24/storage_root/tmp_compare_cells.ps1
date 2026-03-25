param([string]$TemplatePath,[string]$OutputPath,[string]$Sheet,[string[]]$Cells)
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
  $tpl = $excel.Workbooks.Open($TemplatePath,$false,$true)
  $out = $excel.Workbooks.Open($OutputPath,$false,$true)
  $ws1 = $tpl.Worksheets.Item($Sheet)
  $ws2 = $out.Worksheets.Item($Sheet)
  foreach ($cellRef in $Cells) {
    $c1=$ws1.Range($cellRef); $c2=$ws2.Range($cellRef)
    Write-Output ("{0} | T.Text=[{1}] T.Formula=[{2}] | O.Text=[{3}] O.Formula=[{4}]" -f $cellRef,$c1.Text,$c1.Formula,$c2.Text,$c2.Formula)
    [void][Runtime.Interopservices.Marshal]::ReleaseComObject($c1)
    [void][Runtime.Interopservices.Marshal]::ReleaseComObject($c2)
  }
  $tpl.Close($false); $out.Close($false)
  [void][Runtime.Interopservices.Marshal]::ReleaseComObject($ws1)
  [void][Runtime.Interopservices.Marshal]::ReleaseComObject($ws2)
  [void][Runtime.Interopservices.Marshal]::ReleaseComObject($tpl)
  [void][Runtime.Interopservices.Marshal]::ReleaseComObject($out)
}
finally {
  $excel.Quit(); [void][Runtime.Interopservices.Marshal]::ReleaseComObject($excel); [gc]::Collect(); [gc]::WaitForPendingFinalizers()
}
