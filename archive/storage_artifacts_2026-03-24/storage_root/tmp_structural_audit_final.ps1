$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$books = @(
  @{Brand='CHANGAN'; Template='C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\templates\11. Concili. Servicios CHANGAN  2026.xls'; Output='C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_changan_finalaudit2_changan_20260319.xls'; Main='VENTAS'},
  @{Brand='PEUGEOT'; Template='C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\templates\11. Concili. Servicios PEUG  2026.xls'; Output='C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_peug_finalaudit2_peug_20260319.xls'; Main='VENTAS'},
  @{Brand='SUZUKI'; Template='C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\templates\11. Concili. Servicios SZK  2026.xls'; Output='C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_szk_finalaudit2_szk_20260319.xls'; Main='VENTAS'},
  @{Brand='MATRIZ'; Template='C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\templates\11. Concili. Servicios TYT 2026.xls'; Output='C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_tyt_finalaudit_tyt_20260319.xls'; Main='MAY VTAS'}
)
function LastUsed($ws,$order){ try { $c=$ws.Cells.Find('*',$ws.Cells.Item(1,1),-4163,2,$order,2,$false,$false,$false); if($null -eq $c){return 1}; if($order -eq 1){return [int]$c.Row}else{return [int]$c.Column} } catch { return 1 } }
foreach($b in $books){
  $twb=$excel.Workbooks.Open($b.Template,$false,$true)
  $owb=$excel.Workbooks.Open($b.Output,$false,$true)
  try {
    $tNames=@($twb.Worksheets|%{$_.Name}); $oNames=@($owb.Worksheets|%{$_.Name})
    Write-Host ('BRAND=' + $b.Brand + ' SHEET_ORDER=' + ((@($tNames)-join '|') -eq (@($oNames)-join '|')))
    foreach($sheetName in @('REP FACTURACIÓN','NOTA DE CREDITO','PX','REP VTAS','PrecontabilizacionVentas','PrecontabilizacionCostos','PrecontabilizacionCostos (2)','COSTO','ESTADISTICAS',$b.Main)){
      $ts=$twb.Worksheets.Item($sheetName); $os=$owb.Worksheets.Item($sheetName)
      $rows=[Math]::Max((LastUsed $ts 1),(LastUsed $os 1)); $cols=[Math]::Max((LastUsed $ts 2),(LastUsed $os 2));
      $colMis=0; for($c=1;$c -le $cols;$c++){ if([Math]::Abs(([double]$ts.Columns.Item($c).ColumnWidth)-([double]$os.Columns.Item($c).ColumnWidth)) -gt 0.01){$colMis++}}
      $rowMis=0; for($r=1;$r -le $rows;$r++){ if([Math]::Abs(([double]$ts.Rows.Item($r).RowHeight)-([double]$os.Rows.Item($r).RowHeight)) -gt 0.01){$rowMis++}}
      $errs=0; try { $cells=$os.UsedRange.SpecialCells(-4123); foreach($cell in $cells){ $txt=[string]$cell.Text; if($txt -in @('#REF!','#¡REF!','#N/A','#DIV/0!','#VALUE!')){$errs++} } } catch {}
      Write-Host (' ' + $sheetName + ' colMis=' + $colMis + ' rowMis=' + $rowMis + ' errs=' + $errs)
    }
  } finally { $owb.Close($false); $twb.Close($false) }
}
$excel.Quit()
