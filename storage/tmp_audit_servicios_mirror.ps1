$ErrorActionPreference = 'Stop'

function Get-LastUsedCell($worksheet, $searchOrder) {
    try {
        return $worksheet.Cells.Find('*', $worksheet.Cells.Item(1,1), -4163, 2, $searchOrder, 2, $false, $false, $false)
    } catch {
        return $null
    }
}
function Get-LastUsedRow($worksheet) { $cell = Get-LastUsedCell $worksheet 1; if ($null -eq $cell) { return 1 }; return [int]$cell.Row }
function Get-LastUsedCol($worksheet) { $cell = Get-LastUsedCell $worksheet 2; if ($null -eq $cell) { return 1 }; return [int]$cell.Column }
function Get-MatrixValue($matrix, $row, $col, $maxRow, $maxCol) {
    if ($row -gt $maxRow -or $col -gt $maxCol) { return $null }
    if ($matrix -is [System.Array]) { return $matrix.GetValue($row, $col) }
    if ($row -eq 1 -and $col -eq 1) { return $matrix }
    return $null
}
function Normalize-Scalar($value) {
    if ($null -eq $value) { return '' }
    if ($value -is [double] -or $value -is [float] -or $value -is [decimal] -or $value -is [int] -or $value -is [long]) { return [double]$value }
    return ([string]$value).Trim()
}
function ValuesEqual($left, $right) {
    $leftNorm = Normalize-Scalar $left; $rightNorm = Normalize-Scalar $right
    if ($leftNorm -is [double] -and $rightNorm -is [double]) { return [Math]::Abs($leftNorm - $rightNorm) -lt 0.005 }
    if ([string]::IsNullOrWhiteSpace([string]$leftNorm) -and [string]::IsNullOrWhiteSpace([string]$rightNorm)) { return $true }
    return ([string]$leftNorm) -eq ([string]$rightNorm)
}
function Stringify($value) {
    if ($null -eq $value) { return '' }
    if ($value -is [double] -or $value -is [float] -or $value -is [decimal]) { return ('{0:0.###############}' -f [double]$value) }
    return ([string]$value).Trim()
}
function CountFormulaErrors($worksheet) {
    $count = 0
    try {
        $formulaCells = $worksheet.UsedRange.SpecialCells(-4123)
        foreach ($cell in $formulaCells) {
            if (([string]$cell.Formula) -like '*#REF!*' -or ([string]$cell.Text) -eq '#REF!') { $count += 1 }
        }
    } catch {}
    return $count
}
function Compare-Worksheet($templateSheet, $outputSheet) {
    $tLastRow = Get-LastUsedRow $templateSheet; $tLastCol = Get-LastUsedCol $templateSheet
    $oLastRow = Get-LastUsedRow $outputSheet; $oLastCol = Get-LastUsedCol $outputSheet
    $tRange = $templateSheet.Range($templateSheet.Cells.Item(1,1), $templateSheet.Cells.Item($tLastRow, $tLastCol))
    $oRange = $outputSheet.Range($outputSheet.Cells.Item(1,1), $outputSheet.Cells.Item($oLastRow, $oLastCol))
    $tValues = $tRange.Value2; $oValues = $oRange.Value2
    $tFormulas = $tRange.Formula; $oFormulas = $oRange.Formula
    $maxRows = [Math]::Max($tLastRow, $oLastRow); $maxCols = [Math]::Max($tLastCol, $oLastCol)
    $formulaMismatchCount = 0; $valueMismatchCount = 0
    $formulaExamples = New-Object System.Collections.Generic.List[object]
    $valueExamples = New-Object System.Collections.Generic.List[object]
    for ($r = 1; $r -le $maxRows; $r++) {
        for ($c = 1; $c -le $maxCols; $c++) {
            $tFormula = Stringify (Get-MatrixValue $tFormulas $r $c $tLastRow $tLastCol)
            $oFormula = Stringify (Get-MatrixValue $oFormulas $r $c $oLastRow $oLastCol)
            $tValue = Get-MatrixValue $tValues $r $c $tLastRow $tLastCol
            $oValue = Get-MatrixValue $oValues $r $c $oLastRow $oLastCol
            $hasTemplateFormula = $tFormula.StartsWith('='); $hasOutputFormula = $oFormula.StartsWith('=')
            if ($hasTemplateFormula -or $hasOutputFormula) {
                if ($tFormula -ne $oFormula) {
                    $formulaMismatchCount += 1
                    if ($formulaExamples.Count -lt 5) {
                        $formulaExamples.Add([pscustomobject]@{ Cell = $outputSheet.Cells.Item($r,$c).Address($false,$false); TemplateFormula = $tFormula; OutputFormula = $oFormula })
                    }
                }
                continue
            }
            $templateHasValue = -not [string]::IsNullOrWhiteSpace((Stringify $tValue))
            $outputHasValue = -not [string]::IsNullOrWhiteSpace((Stringify $oValue))
            if (-not $templateHasValue -and -not $outputHasValue) { continue }
            if (-not (ValuesEqual $tValue $oValue)) {
                $valueMismatchCount += 1
                if ($valueExamples.Count -lt 5) {
                    $valueExamples.Add([pscustomobject]@{ Cell = $outputSheet.Cells.Item($r,$c).Address($false,$false); TemplateValue = (Stringify $tValue); OutputValue = (Stringify $oValue) })
                }
            }
        }
    }
    $colMismatchCount = 0
    for ($c = 1; $c -le $maxCols; $c++) {
        if ([Math]::Abs(([double]$templateSheet.Columns.Item($c).ColumnWidth) - ([double]$outputSheet.Columns.Item($c).ColumnWidth)) -gt 0.01) { $colMismatchCount += 1 }
    }
    $rowMismatchCount = 0
    for ($r = 1; $r -le $maxRows; $r++) {
        if ([Math]::Abs(([double]$templateSheet.Rows.Item($r).RowHeight) - ([double]$outputSheet.Rows.Item($r).RowHeight)) -gt 0.01) { $rowMismatchCount += 1 }
    }
    $pageSetupMismatch = @()
    foreach ($field in @('PrintArea','Orientation','Zoom','FitToPagesWide','FitToPagesTall','LeftMargin','RightMargin','TopMargin','BottomMargin')) {
        $tv = Stringify ($templateSheet.PageSetup.$field); $ov = Stringify ($outputSheet.PageSetup.$field)
        if ($tv -ne $ov) { $pageSetupMismatch += $field }
    }
    return [pscustomobject]@{
        Sheet = $outputSheet.Name
        TemplateLastRow = $tLastRow
        TemplateLastCol = $tLastCol
        OutputLastRow = $oLastRow
        OutputLastCol = $oLastCol
        FormulaMismatchCount = $formulaMismatchCount
        ValueMismatchCount = $valueMismatchCount
        FormulaExamples = $formulaExamples
        ValueExamples = $valueExamples
        ColumnWidthMismatchCount = $colMismatchCount
        RowHeightMismatchCount = $rowMismatchCount
        PageSetupMismatchFields = $pageSetupMismatch
        RefErrorCount = CountFormulaErrors $outputSheet
    }
}

$brands = @(
    [pscustomobject]@{ Brand='CHANGAN'; Template='C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\templates\11. Concili. Servicios CHANGAN  2026.xls'; Output='C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_changan_mirrorall_20260319.xls'; MainSheet='VENTAS' },
    [pscustomobject]@{ Brand='PEUGEOT'; Template='C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\templates\11. Concili. Servicios PEUG  2026.xls'; Output='C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_peug_mirrorall_20260319.xls'; MainSheet='VENTAS' },
    [pscustomobject]@{ Brand='SUZUKI'; Template='C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\templates\11. Concili. Servicios SZK  2026.xls'; Output='C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_szk_mirrorall_20260319.xls'; MainSheet='VENTAS' },
    [pscustomobject]@{ Brand='MATRIZ'; Template='C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\templates\11. Concili. Servicios TYT 2026.xls'; Output='C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_tyt_mirrorall_20260319.xls'; MainSheet='MAY VTAS' }
)
$keySheetsBase = @('REP FACTURACIÓN','NOTA DE CREDITO','PX','REP VTAS','PrecontabilizacionVentas','PrecontabilizacionCostos','PrecontabilizacionCostos (2)','COSTO','ESTADISTICAS')
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.ScreenUpdating = $false; $excel.EnableEvents = $false
$results = New-Object System.Collections.Generic.List[object]
try {
    foreach ($brand in $brands) {
        $template = $excel.Workbooks.Open($brand.Template, $false, $true)
        $output = $excel.Workbooks.Open($brand.Output, $false, $true)
        try {
            $templateSheetNames = @($template.Worksheets | ForEach-Object { $_.Name })
            $outputSheetNames = @($output.Worksheets | ForEach-Object { $_.Name })
            $sheetOrderMatch = ((@($templateSheetNames) -join '|') -eq (@($outputSheetNames) -join '|'))
            $sheets = @($keySheetsBase + $brand.MainSheet | Select-Object -Unique)
            $sheetResults = New-Object System.Collections.Generic.List[object]
            foreach ($sheetName in $sheets) {
                $tSheet = $template.Worksheets.Item($sheetName)
                $oSheet = $output.Worksheets.Item($sheetName)
                $sheetResults.Add((Compare-Worksheet $tSheet $oSheet))
            }
            $pxSheet = $output.Worksheets.Item('PX')
            $precontSheet = $output.Worksheets.Item('PrecontabilizacionVentas')
            $results.Add([pscustomobject]@{
                Brand = $brand.Brand
                Template = $brand.Template
                Output = $brand.Output
                SheetOrderMatch = $sheetOrderMatch
                TemplateSheetNames = $templateSheetNames
                OutputSheetNames = $outputSheetNames
                PxD5 = Stringify $pxSheet.Range('D5').Value2
                PxH3 = Stringify $pxSheet.Range('H3').Text
                PrecontV7 = Stringify $precontSheet.Range('V7').Text
                PrecontV10 = Stringify $precontSheet.Range('V10').Text
                Sheets = $sheetResults
            })
        } finally {
            $output.Close($false)
            $template.Close($false)
        }
    }
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
$results | ConvertTo-Json -Depth 8 | Set-Content -Path 'C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\audit_servicios_mirrorall_20260319.json' -Encoding UTF8
$results | ConvertTo-Json -Depth 8
