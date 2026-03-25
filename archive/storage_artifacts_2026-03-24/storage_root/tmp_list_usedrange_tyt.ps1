$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open('C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\templates\11. Concili. Servicios TYT 2026.xls', $null, $true)
try {
  foreach($ws in @($wb.Worksheets)) {
    $ur = $ws.UsedRange
    $lastRow = [int]($ur.Row + $ur.Rows.Count - 1)
    $lastCol = [int]($ur.Column + $ur.Columns.Count - 1)
    '{0}|rows={1}|cols={2}|lastRow={3}|lastCol={4}' -f $ws.Name, $ur.Rows.Count, $ur.Columns.Count, $lastRow, $lastCol
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ur)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ws)
  }
} finally {
  $wb.Close($false)
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
  $excel.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
[GC]::Collect()
