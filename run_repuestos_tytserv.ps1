[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$InputTyt,

    [Parameter(Mandatory = $true)]
    [string]$InputPeug,

    [Parameter(Mandatory = $true)]
    [string]$InputChgn,

    [Parameter(Mandatory = $true)]
    [string]$InputSzk,

    [Parameter(Mandatory = $true)]
    [string]$TemplatePath,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'

function Normalize-Text {
    param([object]$Value)

    if ($null -eq $Value) {
        return ''
    }

    return ([string]$Value).Trim()
}

function Resolve-RequiredPath {
    param(
        [string]$Path,
        [string]$Label
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        throw "Parametro vacio: ${Label}"
    }

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "No existe ${Label}: ${Path}"
    }

    return (Resolve-Path -LiteralPath $Path).Path
}

function Get-Worksheet-Safe {
    param(
        [object]$Workbook,
        [string[]]$CandidateNames
    )

    foreach ($candidate in $CandidateNames) {
        $candidateName = (Normalize-Text $candidate).ToUpperInvariant()
        foreach ($worksheet in $Workbook.Worksheets) {
            if ((Normalize-Text $worksheet.Name).ToUpperInvariant() -eq $candidateName) {
                return $worksheet
            }
        }
    }

    throw "No se encontro la hoja requerida: $($CandidateNames -join ', ')"
}

function Get-Used-LastRow {
    param([object]$Worksheet)

    $usedRange = $Worksheet.UsedRange
    try {
        return [Math]::Max(1, [int]$usedRange.Row + [int]$usedRange.Rows.Count - 1)
    }
    finally {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($usedRange)
    }
}

function Find-Row-Containing {
    param(
        [object]$Worksheet,
        [int]$Column,
        [string]$Needle,
        [int]$StartRow = 1,
        [int]$EndRow = 500
    )

    $needleText = (Normalize-Text $Needle).ToUpperInvariant()
    if ($needleText -eq '') {
        return $null
    }

    for ($row = $StartRow; $row -le $EndRow; $row++) {
        $cellText = (Normalize-Text $Worksheet.Cells.Item($row, $Column).Text).ToUpperInvariant()
        if ($cellText -like "*$needleText*") {
            return $row
        }
    }

    return $null
}

function Copy-Rep-Range {
    param(
        [object]$SourceSheet,
        [object]$TargetSheet,
        [int]$LastRow
    )

    $sourceRange = $SourceSheet.Range("A1:AO$LastRow")
    $targetRange = $TargetSheet.Range("A1:AO$LastRow")
    try {
        $null = $sourceRange.Copy($targetRange)
    }
    finally {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($targetRange)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sourceRange)
    }
}

function Clear-Contents-Range {
    param(
        [object]$Worksheet,
        [int]$StartRow,
        [int]$EndRow
    )

    if ($StartRow -gt $EndRow) {
        return
    }

    $range = $Worksheet.Range("A$StartRow:AO$EndRow")
    try {
        $null = $range.ClearContents()
    }
    finally {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($range)
    }
}

function Build-Rep-Lookup {
    param(
        [object]$Worksheet
    )

    $lookup = @{}
    $lastRow = Get-Used-LastRow -Worksheet $Worksheet
    $totalRow = Find-Row-Containing-Anywhere -Worksheet $Worksheet -Needle 'TOTAL GENERAL' -StartRow 1 -EndRow ([Math]::Max(200, $lastRow + 10)) -LastColumn 35
    if ($null -eq $totalRow) {
        $totalRow = $lastRow
    }

    for ($row = 11; $row -lt $totalRow; $row++) {
        $doc = Normalize-Text $Worksheet.Cells.Item($row, 5).Text
        if ($doc -eq '') {
            continue
        }

        if (-not $lookup.ContainsKey($doc)) {
            $lookup[$doc] = @{
                Ruc = Normalize-Text $Worksheet.Cells.Item($row, 9).Text
                ClientCode = Normalize-Text $Worksheet.Cells.Item($row, 10).Text
                ClientName = Normalize-Text $Worksheet.Cells.Item($row, 11).Text
            }
        }
    }

    return $lookup
}

function Test-Looks-MaskedRuc {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $true
    }

    return ($Text.ToUpperInvariant() -like '*X*')
}

function Test-Looks-UnreadableName {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $true
    }

    if ($Text -match '[`~!@#$%^*_=\+\{\}\[\]\|\\<>]') {
        return $true
    }

    $letters = [regex]::Matches($Text, '[A-Za-z]').Count
    return ($letters -lt 4)
}

function Apply-Template-Lookup {
    param(
        [object]$Worksheet,
        [hashtable]$Lookup,
        [int]$StartRow,
        [int]$EndRow
    )

    if ($Lookup.Count -eq 0 -or $StartRow -gt $EndRow) {
        return
    }

    for ($row = $StartRow; $row -le $EndRow; $row++) {
        $doc = Normalize-Text $Worksheet.Cells.Item($row, 5).Text
        if ($doc -eq '' -or -not $Lookup.ContainsKey($doc)) {
            continue
        }

        $reference = $Lookup[$doc]
        $currentRuc = Normalize-Text $Worksheet.Cells.Item($row, 9).Text
        $currentCode = Normalize-Text $Worksheet.Cells.Item($row, 10).Text
        $currentName = Normalize-Text $Worksheet.Cells.Item($row, 11).Text

        $refRuc = Normalize-Text $reference.Ruc
        $refCode = Normalize-Text $reference.ClientCode
        $refName = Normalize-Text $reference.ClientName

        if ($refRuc -ne '' -and $currentRuc -eq '' -and -not (Test-Looks-MaskedRuc $refRuc)) {
            $Worksheet.Cells.Item($row, 9).Value2 = "'" + $refRuc
        }

        if ($refCode -ne '' -and $currentCode -eq '') {
            $Worksheet.Cells.Item($row, 10).Value2 = "'" + $refCode
        }

        if ($refName -ne '' -and $currentName -eq '' -and -not (Test-Looks-UnreadableName $refName)) {
            $Worksheet.Cells.Item($row, 11).Value2 = $refName
        }
    }
}

function Normalize-FormulaText {
    param([object]$Value)

    $text = Normalize-Text $Value
    if ($text -eq '') {
        return ''
    }

    return (($text -replace '\s+', '').ToUpperInvariant())
}

function Assert-Rep-SourceWorksheet {
    param(
        [object]$Worksheet,
        [string]$Label
    )

    $checks = @(
        @{ Row = 9; Column = 5; Needle = '# DOC' },
        @{ Row = 9; Column = 9; Needle = 'RUC' },
        @{ Row = 9; Column = 10; Needle = 'CLIENTE' },
        @{ Row = 9; Column = 11; Needle = 'CLIENTE' },
        @{ Row = 9; Column = 16; Needle = 'ITEM' },
        @{ Row = 9; Column = 18; Needle = 'SUBTOT' }
    )

    foreach ($check in $checks) {
        $actual = (Normalize-Text $Worksheet.Cells.Item($check.Row, $check.Column).Text).ToUpperInvariant()
        $expected = (Normalize-Text $check.Needle).ToUpperInvariant()
        if ($actual -notlike "*$expected*") {
            throw ("La hoja fuente {0} no coincide con la estructura esperada en fila {1} columna {2}. Esperado contiene '{3}' y llego '{4}'." -f $Label, $check.Row, $check.Column, $check.Needle, $actual)
        }
    }
}

function Get-Rep-MayorFormulaMap {
    param(
        [string]$Key,
        [int]$TotalRow
    )

    switch ($Key) {
        'tyt' {
            return @{
                16 = "=+'MY REP TYT'!I38"
                18 = "=+'MY REP TYT'!H53"
            }
        }
        'peug' {
            return @{
                16 = "=+'MY REP PEUG'!I19"
                18 = "=+'MY REP PEUG'!H25"
                25 = ("=+X{0}-'NC REP PEUG'!AD9:AE9" -f $TotalRow)
                26 = ("=X{0}" -f $TotalRow)
            }
        }
        'chgn' {
            return @{
                16 = "=+'MY REP CHGN'!J8"
                19 = "=+'MY REP CHGN'!I15"
            }
        }
        'szk' {
            return @{
                16 = "=+'MY REP SZK'!I31"
                18 = "=+'MY REP SZK'!H50"
            }
        }
        default {
            throw "Clave no soportada: $Key"
        }
    }
}

function Assert-NoExcelErrorsInRange {
    param(
        [object]$Worksheet,
        [int]$StartRow,
        [int]$EndRow,
        [int]$LastColumn,
        [string]$Label
    )

    $errorTexts = @('#DIV/0!', '#N/A', '#NAME?', '#NULL!', '#NUM!', '#REF!', '#VALUE!')
    for ($row = $StartRow; $row -le $EndRow; $row++) {
        for ($column = 1; $column -le $LastColumn; $column++) {
            $text = (Normalize-Text $Worksheet.Cells.Item($row, $column).Text).ToUpperInvariant()
            if ($text -in $errorTexts) {
                throw ("La hoja {0} contiene error de Excel en fila {1} columna {2}: {3}" -f $Label, $row, $column, $text)
            }
        }
    }
}

function Validate-Template-Lookup-Application {
    param(
        [object]$Worksheet,
        [hashtable]$Lookup,
        [int]$StartRow,
        [int]$EndRow,
        [string]$Label
    )

    if ($Lookup.Count -eq 0 -or $StartRow -gt $EndRow) {
        return
    }

    for ($row = $StartRow; $row -le $EndRow; $row++) {
        $doc = Normalize-Text $Worksheet.Cells.Item($row, 5).Text
        if ($doc -eq '' -or -not $Lookup.ContainsKey($doc)) {
            continue
        }

        $reference = $Lookup[$doc]
        $currentRuc = Normalize-Text $Worksheet.Cells.Item($row, 9).Text
        $currentCode = Normalize-Text $Worksheet.Cells.Item($row, 10).Text
        $currentName = Normalize-Text $Worksheet.Cells.Item($row, 11).Text

        if ((Normalize-Text $reference.Ruc) -ne '' -and -not (Test-Looks-MaskedRuc $reference.Ruc) -and $currentRuc -eq '') {
            throw ("La hoja {0} quedo sin RUC para documento {1} en fila {2}." -f $Label, $doc, $row)
        }

        if ((Normalize-Text $reference.ClientCode) -ne '' -and $currentCode -eq '') {
            throw ("La hoja {0} quedo sin codigo cliente para documento {1} en fila {2}." -f $Label, $doc, $row)
        }

        if ((Normalize-Text $reference.ClientName) -ne '' -and -not (Test-Looks-UnreadableName $reference.ClientName) -and $currentName -eq '') {
            throw ("La hoja {0} quedo sin nombre cliente para documento {1} en fila {2}." -f $Label, $doc, $row)
        }
    }
}

function Validate-Rep-Sheet-Output {
    param(
        [object]$Worksheet,
        [object]$TemplateWorksheet,
        [string]$Key,
        [hashtable]$Lookup,
        [int]$ExpectedRows
    )

    $sheetName = Normalize-Text $Worksheet.Name
    $scanEnd = [Math]::Max(200, (Get-Used-LastRow -Worksheet $Worksheet) + 10)
    $totalRow = Find-Row-Containing-Anywhere -Worksheet $Worksheet -Needle 'TOTAL GENERAL' -StartRow 1 -EndRow $scanEnd -LastColumn 35
    $mayorRow = Find-Row-Containing-Anywhere -Worksheet $Worksheet -Needle 'MAYOR' -StartRow 1 -EndRow $scanEnd -LastColumn 35
    $templateMayorRow = Find-Row-Containing-Anywhere -Worksheet $TemplateWorksheet -Needle 'MAYOR' -StartRow 1 -EndRow $scanEnd -LastColumn 35

    if ($null -eq $totalRow) {
        throw "La hoja $sheetName no contiene la fila TOTAL GENERAL."
    }

    if ($null -eq $mayorRow) {
        throw "La hoja $sheetName no contiene la fila MAYOR."
    }

    if ($mayorRow -ne ($totalRow + 1)) {
        throw ("La hoja {0} tiene TOTAL GENERAL en fila {1} y MAYOR en fila {2}; se esperaba MAYOR en fila {3}." -f $sheetName, $totalRow, $mayorRow, ($totalRow + 1))
    }

    if ($null -eq $templateMayorRow) {
        throw "La plantilla de $sheetName no contiene la fila MAYOR."
    }

    $actualRows = [Math]::Max(0, $totalRow - 10)
    if ($actualRows -ne $ExpectedRows) {
        throw ("La hoja {0} quedo con {1} filas REP y se esperaban {2}." -f $sheetName, $actualRows, $ExpectedRows)
    }

    if ((Normalize-Text $Worksheet.Cells.Item($mayorRow, 14).Text).ToUpperInvariant() -ne 'MAYOR') {
        throw "La hoja $sheetName no conserva la etiqueta MAYOR en la fila final."
    }

    $formulaMap = Get-Rep-MayorFormulaMap -Key $Key -TotalRow $totalRow
    $formatColumns = @(14) + @($formulaMap.Keys | ForEach-Object { [int]$_ })

    foreach ($column in $formulaMap.Keys) {
        $expectedFormula = Normalize-FormulaText $formulaMap[$column]
        $actualFormula = Normalize-FormulaText $Worksheet.Cells.Item($mayorRow, [int]$column).Formula
        if ($actualFormula -ne $expectedFormula) {
            throw ("La hoja {0} tiene formula incorrecta en fila {1} columna {2}. Esperado '{3}' y llego '{4}'." -f $sheetName, $mayorRow, $column, $formulaMap[$column], $Worksheet.Cells.Item($mayorRow, [int]$column).Formula)
        }
    }

    foreach ($column in $formatColumns) {
        $outputFormat = Normalize-Text $Worksheet.Cells.Item($mayorRow, [int]$column).NumberFormat
        $templateFormat = Normalize-Text $TemplateWorksheet.Cells.Item($templateMayorRow, [int]$column).NumberFormat
        if ($outputFormat -ne $templateFormat) {
            throw ("La hoja {0} perdio el formato de la fila MAYOR en columna {1}. Esperado '{2}' y llego '{3}'." -f $sheetName, $column, $templateFormat, $outputFormat)
        }
    }

    Validate-Template-Lookup-Application -Worksheet $Worksheet -Lookup $Lookup -StartRow 11 -EndRow ($totalRow - 1) -Label $sheetName
    Assert-NoExcelErrorsInRange -Worksheet $Worksheet -StartRow 1 -EndRow $mayorRow -LastColumn 35 -Label $sheetName
}

function Validate-My-Sheet-Output {
    param(
        [object]$OutputWorkbook,
        [object]$TemplateWorkbook,
        [string]$Key,
        [object]$RepGroups
    )

    $layout = Get-My-Layout -Key $Key
    $outputWorksheet = $null
    $templateWorksheet = $null
    try {
        $outputWorksheet = Get-Worksheet-Safe -Workbook $OutputWorkbook -CandidateNames @($layout.MySheetName)
        $templateWorksheet = Get-Worksheet-Safe -Workbook $TemplateWorkbook -CandidateNames @($layout.MySheetName)

        foreach ($section in $layout.Sections) {
            Assert-NoExcelErrorsInRange -Worksheet $outputWorksheet -StartRow $section.StartRow -EndRow $section.EndRow -LastColumn $layout.SaldoColumn -Label $layout.MySheetName
            foreach ($column in @($layout.DateColumn, $layout.DebitColumn, $layout.CreditColumn, $layout.SaldoColumn)) {
                for ($row = $section.StartRow; $row -le $section.EndRow; $row++) {
                    $outputFormat = Normalize-Text $outputWorksheet.Cells.Item($row, $column).NumberFormat
                    $templateFormat = Normalize-Text $templateWorksheet.Cells.Item($row, $column).NumberFormat
                    if ($outputFormat -ne $templateFormat) {
                        throw ("La hoja {0} perdio formato en fila {1} columna {2}. Esperado '{3}' y llego '{4}'." -f $layout.MySheetName, $row, $column, $templateFormat, $outputFormat)
                    }
                }
            }

            $expectedGroups = switch ((Normalize-Text $section.Name).ToLowerInvariant()) {
                'sales' { $RepGroups.Sales }
                'discount' { $RepGroups.Discount }
                default { @{} }
            }
            $actualGroups = Get-Output-Section-Groups `
                -Worksheet $outputWorksheet `
                -Layout $layout `
                -Section $section

            Assert-Grouped-AmountsMatch `
                -ExpectedGroups $expectedGroups `
                -ActualGroups $actualGroups `
                -Label $layout.MySheetName `
                -SectionName (Normalize-Text $section.Name)
        }
    }
    finally {
        if ($null -ne $templateWorksheet) {
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($templateWorksheet)
        }
        if ($null -ne $outputWorksheet) {
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($outputWorksheet)
        }
    }
}

function Validate-Mayor-Iva-Output {
    param(
        [object]$OutputWorkbook,
        [object]$TemplateWorkbook,
        [hashtable]$RepGroupsByKey
    )

    $outputWorksheet = $null
    $templateWorksheet = $null
    try {
        $outputWorksheet = Get-Worksheet-Safe -Workbook $OutputWorkbook -CandidateNames @('MAYOR IVA')
        $templateWorksheet = Get-Worksheet-Safe -Workbook $TemplateWorkbook -CandidateNames @('MAYOR IVA')

        Assert-NoExcelErrorsInRange -Worksheet $outputWorksheet -StartRow 299 -EndRow 366 -LastColumn 10 -Label 'MAYOR IVA'

        foreach ($column in @(4, 8, 9, 10)) {
            for ($row = 299; $row -le 366; $row++) {
                $outputFormat = Normalize-Text $outputWorksheet.Cells.Item($row, $column).NumberFormat
                $templateFormat = Normalize-Text $templateWorksheet.Cells.Item($row, $column).NumberFormat
                if ($outputFormat -ne $templateFormat) {
                    throw ("La hoja MAYOR IVA perdio formato en fila {0} columna {1}. Esperado '{2}' y llego '{3}'." -f $row, $column, $templateFormat, $outputFormat)
                }
            }
        }

        $expectedGroups = @{}
        foreach ($key in $RepGroupsByKey.Keys) {
            foreach ($groupKey in $RepGroupsByKey[$key].Vat.Keys) {
                $expectedGroups[$groupKey] = $RepGroupsByKey[$key].Vat[$groupKey]
            }
        }

        $actualGroups = Get-Output-Mayor-Iva-Groups -Worksheet $outputWorksheet
        Assert-Grouped-AmountsMatch `
            -ExpectedGroups $expectedGroups `
            -ActualGroups $actualGroups `
            -Label 'MAYOR IVA' `
            -SectionName 'vat'
    }
    finally {
        if ($null -ne $templateWorksheet) {
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($templateWorksheet)
        }
        if ($null -ne $outputWorksheet) {
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($outputWorksheet)
        }
    }
}

function Apply-Mayor-Row {
    param(
        [object]$Worksheet,
        [string]$Key,
        [int]$TotalRow,
        [Nullable[int]]$FormatRow = $null
    )

    $mayorRow = $TotalRow + 1
    $formulaMap = Get-Rep-MayorFormulaMap -Key $Key -TotalRow $TotalRow
    $formatSourceRow = if ($FormatRow.HasValue -and $FormatRow.Value -gt 0) { $FormatRow.Value } else { $mayorRow }
    $formatMap = @{}
    foreach ($column in (@(14) + @($formulaMap.Keys | ForEach-Object { [int]$_ }))) {
        $formatMap[[int]$column] = [string]$Worksheet.Cells.Item($formatSourceRow, [int]$column).NumberFormat
    }

    if ($FormatRow.HasValue -and $FormatRow.Value -gt 0 -and $FormatRow.Value -ne $mayorRow) {
        $formatSource = $Worksheet.Range("A$($FormatRow.Value):AO$($FormatRow.Value)")
        $formatTarget = $Worksheet.Range("A$mayorRow:AO$mayorRow")
        try {
            $null = $formatSource.Copy()
            $null = $formatTarget.PasteSpecial(-4122)
            $Worksheet.Application.CutCopyMode = $false
        }
        finally {
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($formatTarget)
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($formatSource)
        }
    }

    $rowRange = $Worksheet.Range("A$mayorRow:AO$mayorRow")
    try {
        $null = $rowRange.ClearContents()
    }
    finally {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($rowRange)
    }

    $Worksheet.Cells.Item($mayorRow, 14).Value2 = 'MAYOR'

    foreach ($column in $formulaMap.Keys) {
        $Worksheet.Cells.Item($mayorRow, [int]$column).Formula = $formulaMap[$column]
    }

    foreach ($column in $formatMap.Keys) {
        if (-not [string]::IsNullOrWhiteSpace($formatMap[$column])) {
            $Worksheet.Cells.Item($mayorRow, [int]$column).NumberFormat = $formatMap[$column]
        }
    }

    return $mayorRow
}

function Get-Cell-Text {
    param(
        [object]$Worksheet,
        [int]$Row,
        [int]$Column
    )

    return Normalize-Text $Worksheet.Cells.Item($Row, $Column).Text
}

function Get-Cell-Number {
    param(
        [object]$Worksheet,
        [int]$Row,
        [int]$Column
    )

    $value = $Worksheet.Cells.Item($Row, $Column).Value2
    if ($null -eq $value -or $value -eq '') {
        return 0.0
    }

    try {
        return [double]$value
    }
    catch {
        $text = (Normalize-Text $Worksheet.Cells.Item($Row, $Column).Text) -replace ',', ''
        if ($text -eq '') {
            return 0.0
        }

        return [double]::Parse($text, [System.Globalization.CultureInfo]::InvariantCulture)
    }
}

function Get-Rep-DetailName {
    param(
        [string]$Key,
        [string]$Agency
    )

    switch ($Key) {
        'tyt' {
            return 'MOD. REPUESTOS REP01'
        }
        'peug' {
            return 'MOD. REPUESTOS REP06'
        }
        'chgn' {
            return 'MOD. REPUESTOS REP05'
        }
        'szk' {
            if ((Normalize-Text $Agency) -eq '08') {
                return 'MOD. REPUESTOS REP08'
            }

            return 'MOD. REPUESTOS REP07'
        }
        default {
            throw "Clave no soportada para detalle REP: $Key"
        }
    }
}

function Get-Posting-Account {
    param(
        [string]$Key,
        [string]$Category,
        [string]$Form
    )

    $normalizedForm = (Normalize-Text $Form).ToUpperInvariant()

    switch ($Key) {
        'tyt' {
            switch ($Category) {
                'sales' {
                    if ($normalizedForm -eq 'CONTADO') { return '04.01.01.01.0001' }
                    if ($normalizedForm -eq 'CREDITO') { return '04.01.01.01.0003' }
                }
                'discount' {
                    if ($normalizedForm -eq 'CONTADO') { return '04.01.01.01.0005' }
                    if ($normalizedForm -eq 'CREDITO') { return '04.01.01.01.0007' }
                }
            }
        }
        'peug' {
            switch ($Category) {
                'sales' {
                    if ($normalizedForm -eq 'CONTADO') { return '04.01.01.03.0001' }
                    if ($normalizedForm -eq 'CREDITO') { return '04.01.01.03.0003' }
                }
                'discount' {
                    if ($normalizedForm -eq 'CONTADO') { return '04.01.01.03.0005' }
                    if ($normalizedForm -eq 'CREDITO') { return '04.01.01.03.0007' }
                }
            }
        }
        'chgn' {
            switch ($Category) {
                'sales' {
                    if ($normalizedForm -eq 'CONTADO') { return '04.01.01.02.0001' }
                    if ($normalizedForm -eq 'CREDITO') { return '04.01.01.02.0003' }
                }
                'discount' {
                    if ($normalizedForm -eq 'CONTADO') { return '04.01.01.02.0005' }
                }
            }
        }
        'szk' {
            switch ($Category) {
                'sales' {
                    if ($normalizedForm -eq 'CONTADO') { return '04.01.01.04.0001' }
                    if ($normalizedForm -eq 'CREDITO') { return '04.01.01.04.0003' }
                }
                'discount' {
                    if ($normalizedForm -eq 'CONTADO') { return '04.01.01.04.0005' }
                    if ($normalizedForm -eq 'CREDITO') { return '04.01.01.04.0007' }
                }
            }
        }
        default {
            throw "Clave no soportada para cuentas contables: $Key"
        }
    }

    return ''
}

function Add-Grouped-Amount {
    param(
        [hashtable]$Groups,
        [string]$GroupKey,
        [double]$Amount,
        [object]$DateValue,
        [string]$DateText,
        [string]$Seat,
        [string]$Detail
    )

    if ([Math]::Abs($Amount) -lt 0.0000001) {
        return
    }

    if (-not $Groups.ContainsKey($GroupKey)) {
        $Groups[$GroupKey] = @{
            Amount = 0.0
            DateValue = $DateValue
            DateText = $DateText
            Seat = $Seat
            Detail = $Detail
        }
    }

    $entry = $Groups[$GroupKey]
    $entry.Amount = [Math]::Round(([double]$entry.Amount + [double]$Amount), 6)

    if (($null -eq $entry.DateValue -or $entry.DateValue -eq '') -and $null -ne $DateValue -and $DateValue -ne '') {
        $entry.DateValue = $DateValue
    }

    if ((Normalize-Text $entry.DateText) -eq '' -and $DateText -ne '') {
        $entry.DateText = $DateText
    }
}

function Convert-ToExcel-Date {
    param(
        [object]$Value,
        [string]$FallbackText = ''
    )

    if ($null -ne $Value -and $Value -isnot [string]) {
        return $Value
    }

    $text = Normalize-Text $Value
    if ($text -eq '') {
        $text = Normalize-Text $FallbackText
    }
    if ($text -eq '') {
        return $null
    }

    $cultures = @(
        [System.Globalization.CultureInfo]::InvariantCulture,
        [System.Globalization.CultureInfo]::GetCultureInfo('es-EC'),
        [System.Globalization.CultureInfo]::GetCultureInfo('en-US')
    )
    $formats = @(
        'dd/MM/yyyy',
        'd/M/yyyy',
        'dd-MMM-yy',
        'd-MMM-yy',
        'dd-MMM-yyyy',
        'd-MMM-yyyy'
    )

    foreach ($culture in $cultures) {
        foreach ($format in $formats) {
            $parsed = [datetime]::MinValue
            if ([datetime]::TryParseExact($text, $format, $culture, [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) {
                return $parsed.ToOADate()
            }
        }

        $parsed = [datetime]::MinValue
        if ([datetime]::TryParse($text, $culture, [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) {
            return $parsed.ToOADate()
        }
    }

    return $text
}

function Set-Date-CellValue {
    param(
        [object]$Worksheet,
        [int]$Row,
        [int]$Column,
        [object]$Value,
        [string]$FallbackText = ''
    )

    $resolved = Convert-ToExcel-Date -Value $Value -FallbackText $FallbackText
    if ($null -eq $resolved -or $resolved -eq '') {
        $Worksheet.Cells.Item($Row, $Column).ClearContents()
        return
    }

    if ($resolved -is [double] -or $resolved -is [int] -or $resolved -is [decimal]) {
        $Worksheet.Cells.Item($Row, $Column).Value2 = [double]$resolved
        return
    }

    $Worksheet.Cells.Item($Row, $Column).Value2 = [string]$resolved
}

function Find-Row-Containing-Anywhere {
    param(
        [object]$Worksheet,
        [string]$Needle,
        [int]$StartRow = 1,
        [int]$EndRow = 200,
        [int]$LastColumn = 40
    )

    $needleText = (Normalize-Text $Needle).ToUpperInvariant()
    if ($needleText -eq '') {
        return $null
    }

    for ($row = $StartRow; $row -le $EndRow; $row++) {
        for ($column = 1; $column -le $LastColumn; $column++) {
            $cellText = (Normalize-Text $Worksheet.Cells.Item($row, $column).Text).ToUpperInvariant()
            if ($cellText -like "*$needleText*") {
                return $row
            }
        }
    }

    return $null
}

function Build-Rep-Posting-Groups {
    param(
        [object]$Worksheet,
        [string]$Key
    )

    $salesGroups = @{}
    $discountGroups = @{}
    $vatGroups = @{}

    $lastRow = Get-Used-LastRow -Worksheet $Worksheet
    $totalRow = Find-Row-Containing -Worksheet $Worksheet -Column 12 -Needle 'TOTAL GENERAL' -StartRow 1 -EndRow ([Math]::Max(200, $lastRow + 10))
    if ($null -eq $totalRow) {
        $totalRow = $lastRow
    }

    for ($row = 11; $row -lt $totalRow; $row++) {
        $seat = Get-Cell-Text -Worksheet $Worksheet -Row $row -Column 38
        if ($seat -eq '' -or $seat -eq 'ASIENTO') {
            continue
        }

        $form = (Get-Cell-Text -Worksheet $Worksheet -Row $row -Column 40).ToUpperInvariant()
        if ($form -ne 'CONTADO' -and $form -ne 'CREDITO') {
            continue
        }

        $agency = Get-Cell-Text -Worksheet $Worksheet -Row $row -Column 39
        $detail = Get-Rep-DetailName -Key $Key -Agency $agency
        $dateValue = $Worksheet.Cells.Item($row, 3).Value2
        $dateText = Get-Cell-Text -Worksheet $Worksheet -Row $row -Column 3

        $salesAccount = Get-Posting-Account -Key $Key -Category 'sales' -Form $form
        $discountAccount = Get-Posting-Account -Key $Key -Category 'discount' -Form $form

        $salesAmount = Get-Cell-Number -Worksheet $Worksheet -Row $row -Column 18
        $discountAmount = Get-Cell-Number -Worksheet $Worksheet -Row $row -Column 20
        $vatAmount = Get-Cell-Number -Worksheet $Worksheet -Row $row -Column 26

        if ($salesAccount -ne '') {
            Add-Grouped-Amount `
                -Groups $salesGroups `
                -GroupKey ($salesAccount + '|' + $seat + '|' + $detail) `
                -Amount $salesAmount `
                -DateValue $dateValue `
                -DateText $dateText `
                -Seat $seat `
                -Detail $detail
        }

        if ($discountAccount -ne '') {
            Add-Grouped-Amount `
                -Groups $discountGroups `
                -GroupKey ($discountAccount + '|' + $seat + '|' + $detail) `
                -Amount $discountAmount `
                -DateValue $dateValue `
                -DateText $dateText `
                -Seat $seat `
                -Detail $detail
        }

        Add-Grouped-Amount `
            -Groups $vatGroups `
            -GroupKey ($seat + '|' + $detail) `
            -Amount $vatAmount `
            -DateValue $dateValue `
            -DateText $dateText `
            -Seat $seat `
            -Detail $detail
    }

    return @{
        Sales = $salesGroups
        Discount = $discountGroups
        Vat = $vatGroups
    }
}

function Get-My-Layout {
    param([string]$Key)

    switch ($Key) {
        'tyt' {
            return @{
                MySheetName = 'MY REP TYT'
                DetailColumn = 7
                SeatColumn = 6
                DateColumn = 4
                DebitColumn = 8
                CreditColumn = 9
                SaldoColumn = 10
                Sections = @(
                    @{ Name = 'sales'; StartRow = 2; EndRow = 37; AmountColumn = 9; OppositeColumn = 8 },
                    @{ Name = 'discount'; StartRow = 42; EndRow = 52; AmountColumn = 8; OppositeColumn = 9 }
                )
            }
        }
        'peug' {
            return @{
                MySheetName = 'MY REP PEUG'
                DetailColumn = 7
                SeatColumn = 6
                DateColumn = 4
                DebitColumn = 8
                CreditColumn = 9
                SaldoColumn = 10
                Sections = @(
                    @{ Name = 'sales'; StartRow = 2; EndRow = 18; AmountColumn = 9; OppositeColumn = 8 },
                    @{ Name = 'discount'; StartRow = 21; EndRow = 24; AmountColumn = 8; OppositeColumn = 9 }
                )
            }
        }
        'chgn' {
            return @{
                MySheetName = 'MY REP CHGN'
                DetailColumn = 8
                SeatColumn = 6
                DateColumn = 4
                DebitColumn = 9
                CreditColumn = 10
                SaldoColumn = 11
                Sections = @(
                    @{ Name = 'sales'; StartRow = 2; EndRow = 7; AmountColumn = 10; OppositeColumn = 9 },
                    @{ Name = 'discount'; StartRow = 12; EndRow = 14; AmountColumn = 9; OppositeColumn = 10 }
                )
            }
        }
        'szk' {
            return @{
                MySheetName = 'MY REP SZK'
                DetailColumn = 7
                SeatColumn = 6
                DateColumn = 4
                DebitColumn = 8
                CreditColumn = 9
                SaldoColumn = 10
                Sections = @(
                    @{ Name = 'sales'; StartRow = 2; EndRow = 30; AmountColumn = 9; OppositeColumn = 8 },
                    @{ Name = 'discount'; StartRow = 34; EndRow = 49; AmountColumn = 8; OppositeColumn = 9 }
                )
            }
        }
        default {
            throw "Clave no soportada para layout MY: $Key"
        }
    }
}

function Get-Output-Section-Groups {
    param(
        [object]$Worksheet,
        [hashtable]$Layout,
        [hashtable]$Section
    )

    $groups = @{}

    for ($row = $Section.StartRow; $row -le $Section.EndRow; $row++) {
        $amount = Get-Cell-Number -Worksheet $Worksheet -Row $row -Column $Section.AmountColumn
        $opposite = Get-Cell-Number -Worksheet $Worksheet -Row $row -Column $Section.OppositeColumn
        if ([Math]::Abs($amount) -lt 0.0000001 -or [Math]::Abs($opposite) -gt 0.0000001) {
            continue
        }

        $account = Get-Cell-Text -Worksheet $Worksheet -Row $row -Column 1
        $seat = Get-Cell-Text -Worksheet $Worksheet -Row $row -Column $Layout.SeatColumn
        $detail = Get-Cell-Text -Worksheet $Worksheet -Row $row -Column $Layout.DetailColumn
        if ($account -eq '' -or $seat -eq '' -or $detail -eq '') {
            continue
        }

        Add-Grouped-Amount `
            -Groups $groups `
            -GroupKey ($account + '|' + $seat + '|' + $detail) `
            -Amount $amount `
            -DateValue $null `
            -DateText '' `
            -Seat $seat `
            -Detail $detail
    }

    return $groups
}

function Get-Output-Mayor-Iva-Groups {
    param([object]$Worksheet)

    $groups = @{}

    for ($row = 299; $row -le 366; $row++) {
        $type = (Get-Cell-Text -Worksheet $Worksheet -Row $row -Column 5).ToUpperInvariant()
        $seat = Get-Cell-Text -Worksheet $Worksheet -Row $row -Column 6
        $detail = Get-Cell-Text -Worksheet $Worksheet -Row $row -Column 7
        $debit = Get-Cell-Number -Worksheet $Worksheet -Row $row -Column 8
        $credit = Get-Cell-Number -Worksheet $Worksheet -Row $row -Column 9

        if ($type -ne 'REPTO' -or $seat -eq '' -or $detail -eq '') {
            continue
        }

        if ([Math]::Abs($credit) -lt 0.0000001 -or [Math]::Abs($debit) -gt 0.0000001) {
            continue
        }

        Add-Grouped-Amount `
            -Groups $groups `
            -GroupKey ($seat + '|' + $detail) `
            -Amount $credit `
            -DateValue $null `
            -DateText '' `
            -Seat $seat `
            -Detail $detail
    }

    return $groups
}

function Assert-Grouped-AmountsMatch {
    param(
        [hashtable]$ExpectedGroups,
        [hashtable]$ActualGroups,
        [string]$Label,
        [string]$SectionName
    )

    $allKeys = @(
        @($ExpectedGroups.Keys) + @($ActualGroups.Keys) |
            Sort-Object -Unique
    )

    foreach ($groupKey in $allKeys) {
        $expectedAmount = if ($ExpectedGroups.ContainsKey($groupKey)) {
            [double]$ExpectedGroups[$groupKey].Amount
        } else {
            0.0
        }

        $actualAmount = if ($ActualGroups.ContainsKey($groupKey)) {
            [double]$ActualGroups[$groupKey].Amount
        } else {
            0.0
        }

        if ([Math]::Abs($expectedAmount - $actualAmount) -ge 0.01) {
            throw ("La hoja {0} no cuadra en {1} para el grupo {2}. Esperado {3:N2} y llego {4:N2}." -f $Label, $SectionName, $groupKey, $expectedAmount, $actualAmount)
        }
    }
}

function Get-Template-Section-Groups {
    param(
        [object]$Worksheet,
        [hashtable]$Layout,
        [hashtable]$Section
    )

    $groups = @{}

    for ($row = $Section.StartRow; $row -le $Section.EndRow; $row++) {
        $amount = Get-Cell-Number -Worksheet $Worksheet -Row $row -Column $Section.AmountColumn
        $opposite = Get-Cell-Number -Worksheet $Worksheet -Row $row -Column $Section.OppositeColumn
        if ([Math]::Abs($amount) -lt 0.0000001 -or [Math]::Abs($opposite) -gt 0.0000001) {
            continue
        }

        $account = Get-Cell-Text -Worksheet $Worksheet -Row $row -Column 1
        $seat = Get-Cell-Text -Worksheet $Worksheet -Row $row -Column $Layout.SeatColumn
        $detail = Get-Cell-Text -Worksheet $Worksheet -Row $row -Column $Layout.DetailColumn
        if ($account -eq '' -or $seat -eq '' -or $detail -eq '') {
            continue
        }

        $groupKey = $account + '|' + $seat + '|' + $detail
        if (-not $groups.ContainsKey($groupKey)) {
            $groups[$groupKey] = @{
                BaseTotal = 0.0
                Rows = @()
            }
        }

        $groups[$groupKey].BaseTotal = [Math]::Round(([double]$groups[$groupKey].BaseTotal + [double]$amount), 6)
        $groups[$groupKey].Rows += @{
            RowNumber = $row
            BaseAmount = [double]$amount
        }
    }

    return $groups
}

function Get-Section-Opening-Balance {
    param(
        [object]$Worksheet,
        [hashtable]$Layout,
        [hashtable]$Section
    )

    for ($row = $Section.StartRow; $row -le $Section.EndRow; $row++) {
        $account = Get-Cell-Text -Worksheet $Worksheet -Row $row -Column 1
        $detail = Get-Cell-Text -Worksheet $Worksheet -Row $row -Column $Layout.DetailColumn
        $debit = Get-Cell-Number -Worksheet $Worksheet -Row $row -Column $Layout.DebitColumn
        $credit = Get-Cell-Number -Worksheet $Worksheet -Row $row -Column $Layout.CreditColumn
        $saldo = Get-Cell-Number -Worksheet $Worksheet -Row $row -Column $Layout.SaldoColumn

        if ($account -eq '' -and $detail -eq '' -and [Math]::Abs($debit) -lt 0.0000001 -and [Math]::Abs($credit) -lt 0.0000001 -and [Math]::Abs($saldo) -lt 0.0000001) {
            continue
        }

        return [Math]::Round(($saldo - $debit + $credit), 2)
    }

    return 0.0
}

function Apply-Section-Scaling {
    param(
        [object]$OutputWorksheet,
        [object]$TemplateWorksheet,
        [hashtable]$Layout,
        [hashtable]$Section,
        [hashtable]$SourceGroups,
        [string]$Key
)

    $templateGroups = Get-Template-Section-Groups -Worksheet $TemplateWorksheet -Layout $Layout -Section $Section
    $matchedGroups = @{}
    $amountColumn = [int]$Section.AmountColumn
    $oppositeColumn = [int]$Section.OppositeColumn
    $dateColumn = [int]$Layout.DateColumn
    $seatColumn = [int]$Layout.SeatColumn
    $detailColumn = [int]$Layout.DetailColumn

    for ($row = $Section.StartRow; $row -le $Section.EndRow; $row++) {
        $templateAmount = Get-Cell-Number -Worksheet $TemplateWorksheet -Row $row -Column $amountColumn
        $templateOpposite = Get-Cell-Number -Worksheet $TemplateWorksheet -Row $row -Column $oppositeColumn
        if ([Math]::Abs($templateAmount) -lt 0.0000001 -and [Math]::Abs($templateOpposite) -gt 0.0000001) {
            $OutputWorksheet.Cells.Item($row, $amountColumn).Value2 = 0.0
            $OutputWorksheet.Cells.Item($row, $oppositeColumn).Value2 = 0.0
            $OutputWorksheet.Cells.Item($row, $dateColumn).ClearContents()
            $OutputWorksheet.Cells.Item($row, $seatColumn).ClearContents()
            $OutputWorksheet.Cells.Item($row, $detailColumn).ClearContents()
        }
    }

    foreach ($groupKey in $templateGroups.Keys) {
        $templateGroup = $templateGroups[$groupKey]
        $sourceGroup = $null
        $targetTotal = 0.0
        if ($SourceGroups.ContainsKey($groupKey)) {
            $sourceGroup = $SourceGroups[$groupKey]
            $targetTotal = [double]$sourceGroup.Amount
            $matchedGroups[$groupKey] = $true
        }

        $rows = @($templateGroup.Rows)
        $rowCount = $rows.Count
        if ($rowCount -eq 0) {
            continue
        }

        $remaining = [Math]::Round($targetTotal, 6)
        for ($index = 0; $index -lt $rowCount; $index++) {
            $currentRow = $rows[$index]
            $rowNumber = [int]$currentRow.RowNumber

            if ($index -eq ($rowCount - 1)) {
                $scaledAmount = [Math]::Round($remaining, 2)
            }
            elseif ([double]$templateGroup.BaseTotal -ne 0.0) {
                $scaledAmount = [Math]::Round(($targetTotal * ([double]$currentRow.BaseAmount / [double]$templateGroup.BaseTotal)), 2)
                $remaining = [Math]::Round(($remaining - $scaledAmount), 6)
            }
            else {
                $scaledAmount = 0.0
            }

            try {
                $OutputWorksheet.Cells.Item($rowNumber, $amountColumn).Value2 = [double]$scaledAmount
                $OutputWorksheet.Cells.Item($rowNumber, $oppositeColumn).Value2 = 0.0
            }
            catch {
                throw ("No se pudo escribir MY {0} fila {1} columna {2} valor {3}: {4}" -f $Section.Name, $rowNumber, $amountColumn, $scaledAmount, $_.Exception.Message)
            }

            if ($null -ne $sourceGroup) {
                Set-Date-CellValue `
                    -Worksheet $OutputWorksheet `
                    -Row $rowNumber `
                    -Column $dateColumn `
                    -Value $sourceGroup.DateValue `
                    -FallbackText $sourceGroup.DateText

                $OutputWorksheet.Cells.Item($rowNumber, $seatColumn).Value2 = [string]$sourceGroup.Seat
                $OutputWorksheet.Cells.Item($rowNumber, $detailColumn).Value2 = [string]$sourceGroup.Detail
            }
        }
    }

    foreach ($sourceGroupKey in $SourceGroups.Keys) {
        if (-not $matchedGroups.ContainsKey($sourceGroupKey)) {
            Write-Output ("WARN|{0}|my_{1}_missing={2}" -f $Key, $Section.Name, $sourceGroupKey)
        }
    }
}

function Recalculate-Section-Saldo {
    param(
        [object]$OutputWorksheet,
        [object]$TemplateWorksheet,
        [hashtable]$Layout,
        [hashtable]$Section
)

    $saldoColumn = [int]$Layout.SaldoColumn
    $runningBalance = 0.0
    $currentAccount = ''
    $hasBalance = $false

    for ($row = $Section.StartRow; $row -le $Section.EndRow; $row++) {
        $templateAccount = Get-Cell-Text -Worksheet $TemplateWorksheet -Row $row -Column 1
        $templateDetail = Get-Cell-Text -Worksheet $TemplateWorksheet -Row $row -Column $Layout.DetailColumn
        $templateDebit = Get-Cell-Number -Worksheet $TemplateWorksheet -Row $row -Column $Layout.DebitColumn
        $templateCredit = Get-Cell-Number -Worksheet $TemplateWorksheet -Row $row -Column $Layout.CreditColumn
        $templateSaldo = Get-Cell-Number -Worksheet $TemplateWorksheet -Row $row -Column $Layout.SaldoColumn

        if ($templateAccount -eq '' -and $templateDetail -eq '' -and [Math]::Abs($templateDebit) -lt 0.0000001 -and [Math]::Abs($templateCredit) -lt 0.0000001 -and [Math]::Abs($templateSaldo) -lt 0.0000001) {
            $OutputWorksheet.Cells.Item($row, $saldoColumn).ClearContents()
            continue
        }

        if (-not $hasBalance -or $templateAccount -ne $currentAccount) {
            $currentAccount = $templateAccount
            $runningBalance = [Math]::Round(($templateSaldo - $templateDebit + $templateCredit), 2)
            $hasBalance = $true
        }

        $debit = Get-Cell-Number -Worksheet $OutputWorksheet -Row $row -Column $Layout.DebitColumn
        $credit = Get-Cell-Number -Worksheet $OutputWorksheet -Row $row -Column $Layout.CreditColumn
        $runningBalance = [Math]::Round(($runningBalance + $debit - $credit), 2)
        $OutputWorksheet.Cells.Item($row, $saldoColumn).Value2 = [double]$runningBalance
    }
}

function Update-My-Sheet-From-Rep {
    param(
        [object]$OutputWorkbook,
        [object]$TemplateWorkbook,
        [string]$Key,
        [string]$RepSheetName
    )

    $layout = Get-My-Layout -Key $Key
    $repWorksheet = $null
    $outputWorksheet = $null
    $templateWorksheet = $null
    try {
        $repWorksheet = Get-Worksheet-Safe -Workbook $OutputWorkbook -CandidateNames @($RepSheetName)
        $outputWorksheet = Get-Worksheet-Safe -Workbook $OutputWorkbook -CandidateNames @($layout.MySheetName)
        $templateWorksheet = Get-Worksheet-Safe -Workbook $TemplateWorkbook -CandidateNames @($layout.MySheetName)

        $repGroups = Build-Rep-Posting-Groups -Worksheet $repWorksheet -Key $Key

        foreach ($section in $layout.Sections) {
            if ($section.Name -eq 'sales') {
                Apply-Section-Scaling `
                    -OutputWorksheet $outputWorksheet `
                    -TemplateWorksheet $templateWorksheet `
                    -Layout $layout `
                    -Section $section `
                    -SourceGroups $repGroups.Sales `
                    -Key $Key
            }
            elseif ($section.Name -eq 'discount') {
                Apply-Section-Scaling `
                    -OutputWorksheet $outputWorksheet `
                    -TemplateWorksheet $templateWorksheet `
                    -Layout $layout `
                    -Section $section `
                    -SourceGroups $repGroups.Discount `
                    -Key $Key
            }

            Recalculate-Section-Saldo `
                -OutputWorksheet $outputWorksheet `
                -TemplateWorksheet $templateWorksheet `
                -Layout $layout `
                -Section $section
        }

        Clear-My-Devol-Rows `
            -OutputWorkbook $OutputWorkbook `
            -TemplateWorkbook $TemplateWorkbook `
            -Key $Key

        Write-Output ("INFO|{0}|my_sheet={1}" -f $Key, $layout.MySheetName)
        return $repGroups
    }
    finally {
        if ($null -ne $templateWorksheet) {
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($templateWorksheet)
        }
        if ($null -ne $outputWorksheet) {
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($outputWorksheet)
        }
        if ($null -ne $repWorksheet) {
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($repWorksheet)
        }
    }
}

function Clear-My-Devol-Rows {
    param(
        [object]$OutputWorkbook,
        [object]$TemplateWorkbook,
        [string]$Key
    )

    $layout = Get-My-Layout -Key $Key
    $outputWorksheet = $null
    $templateWorksheet = $null
    try {
        $outputWorksheet = Get-Worksheet-Safe -Workbook $OutputWorkbook -CandidateNames @($layout.MySheetName)
        $templateWorksheet = Get-Worksheet-Safe -Workbook $TemplateWorkbook -CandidateNames @($layout.MySheetName)

        $lastSectionEnd = 0
        foreach ($section in $layout.Sections) {
            if ([int]$section.EndRow -gt $lastSectionEnd) {
                $lastSectionEnd = [int]$section.EndRow
            }
        }

        $lastRow = Get-Used-LastRow -Worksheet $templateWorksheet
        for ($row = $lastSectionEnd + 1; $row -le $lastRow; $row++) {
            $account = Get-Cell-Text -Worksheet $templateWorksheet -Row $row -Column 1
            $name = (Get-Cell-Text -Worksheet $templateWorksheet -Row $row -Column 2).ToUpperInvariant()
            if ($account -eq '' -or $name -notlike '*DEVOL*') {
                continue
            }

            $outputWorksheet.Cells.Item($row, $layout.DateColumn).ClearContents()
            $outputWorksheet.Cells.Item($row, $layout.SeatColumn).ClearContents()
            $outputWorksheet.Cells.Item($row, $layout.DetailColumn).ClearContents()
            $outputWorksheet.Cells.Item($row, $layout.DebitColumn).Value2 = 0.0
            $outputWorksheet.Cells.Item($row, $layout.CreditColumn).Value2 = 0.0
            $outputWorksheet.Cells.Item($row, $layout.SaldoColumn).ClearContents()
        }
    }
    finally {
        if ($null -ne $templateWorksheet) {
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($templateWorksheet)
        }
        if ($null -ne $outputWorksheet) {
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($outputWorksheet)
        }
    }
}

function Clear-Nc-Sheet-NoSource {
    param(
        [object]$Workbook,
        [string]$SheetName
    )

    $worksheet = $null
    try {
        $worksheet = Get-Worksheet-Safe -Workbook $Workbook -CandidateNames @($SheetName)
        $totalRow = Find-Row-Containing-Anywhere -Worksheet $worksheet -Needle 'TOTAL GENERAL' -StartRow 1 -EndRow 40 -LastColumn 35
        $mayorRow = Find-Row-Containing-Anywhere -Worksheet $worksheet -Needle 'MAYOR' -StartRow 1 -EndRow 40 -LastColumn 35

        if ($null -eq $totalRow) {
            return
        }

        if ($totalRow -gt 8) {
            $detailRange = $worksheet.Range("A8:AE$($totalRow - 1)")
            try {
                $null = $detailRange.ClearContents()
            }
            finally {
                [void][Runtime.InteropServices.Marshal]::ReleaseComObject($detailRange)
            }
        }

        foreach ($column in 22..24) {
            $worksheet.Cells.Item($totalRow, $column).Value2 = 0.0
        }

        if ($null -ne $mayorRow) {
            $worksheet.Cells.Item($mayorRow, 22).Value2 = 'MAYOR'
        }
    }
    finally {
        if ($null -ne $worksheet) {
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet)
        }
    }
}

function Get-Mayor-Iva-Template-Groups {
    param(
        [object]$Worksheet,
        [int]$StartRow,
        [int]$EndRow
    )

    $groups = @{}

    for ($row = $StartRow; $row -le $EndRow; $row++) {
        $type = (Get-Cell-Text -Worksheet $Worksheet -Row $row -Column 5).ToUpperInvariant()
        $seat = Get-Cell-Text -Worksheet $Worksheet -Row $row -Column 6
        $detail = Get-Cell-Text -Worksheet $Worksheet -Row $row -Column 7
        $debit = Get-Cell-Number -Worksheet $Worksheet -Row $row -Column 8
        $credit = Get-Cell-Number -Worksheet $Worksheet -Row $row -Column 9

        if ($type -ne 'REPTO' -or $seat -eq '' -or $detail -eq '') {
            continue
        }

        if ([Math]::Abs($credit) -lt 0.0000001 -or [Math]::Abs($debit) -gt 0.0000001) {
            continue
        }

        $groupKey = $seat + '|' + $detail
        if (-not $groups.ContainsKey($groupKey)) {
            $groups[$groupKey] = @{
                BaseTotal = 0.0
                Rows = @()
            }
        }

        $groups[$groupKey].BaseTotal = [Math]::Round(([double]$groups[$groupKey].BaseTotal + [double]$credit), 6)
        $groups[$groupKey].Rows += @{
            RowNumber = $row
            BaseAmount = [double]$credit
        }
    }

    return $groups
}

function Get-Mayor-Iva-Opening-Balance {
    param(
        [object]$Worksheet,
        [int]$StartRow
    )

    $saldo = Get-Cell-Number -Worksheet $Worksheet -Row $StartRow -Column 10
    $debit = Get-Cell-Number -Worksheet $Worksheet -Row $StartRow -Column 8
    $credit = Get-Cell-Number -Worksheet $Worksheet -Row $StartRow -Column 9
    return [Math]::Round(($saldo - $debit + $credit), 2)
}

function Update-Mayor-Iva-From-Rep {
    param(
        [object]$OutputWorkbook,
        [object]$TemplateWorkbook,
        [hashtable]$RepGroupsByKey
    )

    $outputWorksheet = $null
    $templateWorksheet = $null
    try {
        $outputWorksheet = Get-Worksheet-Safe -Workbook $OutputWorkbook -CandidateNames @('MAYOR IVA')
        $templateWorksheet = Get-Worksheet-Safe -Workbook $TemplateWorkbook -CandidateNames @('MAYOR IVA')

        $templateGroups = Get-Mayor-Iva-Template-Groups -Worksheet $templateWorksheet -StartRow 299 -EndRow 366
        $sourceGroups = @{}

        foreach ($key in $RepGroupsByKey.Keys) {
            foreach ($groupKey in $RepGroupsByKey[$key].Vat.Keys) {
                $sourceGroups[$groupKey] = $RepGroupsByKey[$key].Vat[$groupKey]
            }
        }

        $matchedGroups = @{}
        foreach ($groupKey in $templateGroups.Keys) {
            $templateGroup = $templateGroups[$groupKey]
            $sourceGroup = $null
            $targetTotal = 0.0
            if ($sourceGroups.ContainsKey($groupKey)) {
                $sourceGroup = $sourceGroups[$groupKey]
                $targetTotal = [double]$sourceGroup.Amount
                $matchedGroups[$groupKey] = $true
            }

            $rows = @($templateGroup.Rows)
            $remaining = [Math]::Round($targetTotal, 6)
            for ($index = 0; $index -lt $rows.Count; $index++) {
                $rowNumber = [int]$rows[$index].RowNumber
                if ($index -eq ($rows.Count - 1)) {
                    $scaledAmount = [Math]::Round($remaining, 2)
                }
                elseif ([double]$templateGroup.BaseTotal -ne 0.0) {
                    $scaledAmount = [Math]::Round(($targetTotal * ([double]$rows[$index].BaseAmount / [double]$templateGroup.BaseTotal)), 2)
                    $remaining = [Math]::Round(($remaining - $scaledAmount), 6)
                }
                else {
                    $scaledAmount = 0.0
                }

                $outputWorksheet.Cells.Item($rowNumber, 8).Value2 = 0.0
                $outputWorksheet.Cells.Item($rowNumber, 9).Value2 = [double]$scaledAmount

                if ($null -ne $sourceGroup) {
                    Set-Date-CellValue `
                        -Worksheet $outputWorksheet `
                        -Row $rowNumber `
                        -Column 4 `
                        -Value $sourceGroup.DateValue `
                        -FallbackText $sourceGroup.DateText

                    $outputWorksheet.Cells.Item($rowNumber, 6).Value2 = [string]$sourceGroup.Seat
                    $outputWorksheet.Cells.Item($rowNumber, 7).Value2 = [string]$sourceGroup.Detail
                }
            }
        }

        foreach ($sourceGroupKey in $sourceGroups.Keys) {
            if (-not $matchedGroups.ContainsKey($sourceGroupKey)) {
                Write-Output ("WARN|mayor_iva|missing={0}" -f $sourceGroupKey)
            }
        }

        $runningBalance = Get-Mayor-Iva-Opening-Balance -Worksheet $templateWorksheet -StartRow 299
        for ($row = 299; $row -le 366; $row++) {
            $templateType = Get-Cell-Text -Worksheet $templateWorksheet -Row $row -Column 5
            $templateSeat = Get-Cell-Text -Worksheet $templateWorksheet -Row $row -Column 6
            $templateDetail = Get-Cell-Text -Worksheet $templateWorksheet -Row $row -Column 7
            $templateDebit = Get-Cell-Number -Worksheet $templateWorksheet -Row $row -Column 8
            $templateCredit = Get-Cell-Number -Worksheet $templateWorksheet -Row $row -Column 9
            $templateSaldo = Get-Cell-Number -Worksheet $templateWorksheet -Row $row -Column 10

            if ($templateType -eq '' -and $templateSeat -eq '' -and $templateDetail -eq '' -and [Math]::Abs($templateDebit) -lt 0.0000001 -and [Math]::Abs($templateCredit) -lt 0.0000001 -and [Math]::Abs($templateSaldo) -lt 0.0000001) {
                $outputWorksheet.Cells.Item($row, 10).ClearContents()
                continue
            }

            $debit = Get-Cell-Number -Worksheet $outputWorksheet -Row $row -Column 8
            $credit = Get-Cell-Number -Worksheet $outputWorksheet -Row $row -Column 9
            $runningBalance = [Math]::Round(($runningBalance + $debit - $credit), 2)
            $outputWorksheet.Cells.Item($row, 10).Value2 = [double]$runningBalance
        }

        Write-Output 'INFO|mayor_iva|updated=1'
    }
    finally {
        if ($null -ne $templateWorksheet) {
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($templateWorksheet)
        }
        if ($null -ne $outputWorksheet) {
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($outputWorksheet)
        }
    }
}

function Process-Rep-Sheet {
    param(
        [object]$Excel,
        [object]$OutputWorkbook,
        [string]$SourcePath,
        [string]$TargetSheetName,
        [string]$Key,
        [string]$Label,
        [hashtable]$Lookup
    )

    $sourceWorkbook = $null
    $sourceSheet = $null
    $targetSheet = $null
    try {
        $sourceWorkbook = $Excel.Workbooks.Open($SourcePath, 0, $true)
        try {
            $sourceSheet = Get-Worksheet-Safe -Workbook $sourceWorkbook -CandidateNames @(
                'RepLibroVentasGeneral',
                $TargetSheetName
            )
        }
        catch {
            $sourceSheet = $sourceWorkbook.Worksheets.Item(1)
        }

        Assert-Rep-SourceWorksheet -Worksheet $sourceSheet -Label $Label

        $targetSheet = Get-Worksheet-Safe -Workbook $OutputWorkbook -CandidateNames @($TargetSheetName)
        $oldUsedLastRow = Get-Used-LastRow -Worksheet $targetSheet
        $oldMayorRow = Find-Row-Containing -Worksheet $targetSheet -Column 14 -Needle 'MAYOR' -StartRow 1 -EndRow ([Math]::Max(200, $oldUsedLastRow + 20))

        $sourceLastRow = Get-Used-LastRow -Worksheet $sourceSheet
        Copy-Rep-Range -SourceSheet $sourceSheet -TargetSheet $targetSheet -LastRow $sourceLastRow

        $scanEnd = [Math]::Max($sourceLastRow + 10, $oldUsedLastRow + 10)
        $newTotalRow = Find-Row-Containing-Anywhere -Worksheet $targetSheet -Needle 'TOTAL GENERAL' -StartRow 1 -EndRow $scanEnd -LastColumn 35
        if ($null -eq $newTotalRow) {
            $newTotalRow = $sourceLastRow
        }

        Apply-Template-Lookup -Worksheet $targetSheet -Lookup $Lookup -StartRow 11 -EndRow ($newTotalRow - 1)
        $newMayorRow = Apply-Mayor-Row -Worksheet $targetSheet -Key $Key -TotalRow $newTotalRow -FormatRow $oldMayorRow

        $cleanupStart = $newMayorRow + 1
        $cleanupEnd = [Math]::Max($oldUsedLastRow, $sourceLastRow + 10)
        if ($cleanupStart -le $cleanupEnd) {
            Clear-Contents-Range -Worksheet $targetSheet -StartRow $cleanupStart -EndRow $cleanupEnd
        }

        $rowCount = [Math]::Max(0, $newTotalRow - 10)
        Write-Output ("INFO|{0}|rows={1}" -f $Key, $rowCount)
        Write-Output ("INFO|{0}|sheet={1}" -f $Key, $TargetSheetName)
        Write-Output ("INFO|{0}|label={1}" -f $Key, $Label)

        return [pscustomobject]@{
            Key = $Key
            Label = $Label
            TargetSheetName = $TargetSheetName
            RowCount = $rowCount
        }
    }
    finally {
        if ($null -ne $sourceWorkbook) {
            $sourceWorkbook.Close($false)
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sourceWorkbook)
        }
        if ($null -ne $targetSheet) {
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($targetSheet)
        }
        if ($null -ne $sourceSheet) {
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sourceSheet)
        }
    }
}

$resolvedInputTyt = Resolve-RequiredPath -Path $InputTyt -Label 'InputTyt'
$resolvedInputPeug = Resolve-RequiredPath -Path $InputPeug -Label 'InputPeug'
$resolvedInputChgn = Resolve-RequiredPath -Path $InputChgn -Label 'InputChgn'
$resolvedInputSzk = Resolve-RequiredPath -Path $InputSzk -Label 'InputSzk'
$resolvedTemplatePath = Resolve-RequiredPath -Path $TemplatePath -Label 'TemplatePath'

$outputDirectory = Split-Path -Path $OutputPath -Parent
if ([string]::IsNullOrWhiteSpace($outputDirectory)) {
    throw 'No se pudo resolver la carpeta de salida.'
}
if (-not (Test-Path -LiteralPath $outputDirectory)) {
    $null = New-Item -ItemType Directory -Path $outputDirectory -Force
}

$resolvedOutputPath = [System.IO.Path]::GetFullPath($OutputPath)
Copy-Item -LiteralPath $resolvedTemplatePath -Destination $resolvedOutputPath -Force

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false
$excel.EnableEvents = $false
$excel.AskToUpdateLinks = $false
try {
    $excel.Calculation = -4135
}
catch {
}

$outputWorkbook = $null
$templateWorkbook = $null
$lookupTytSheet = $null
$lookupPeugSheet = $null
$lookupChgnSheet = $null
$lookupSzkSheet = $null
try {
    $outputWorkbook = $excel.Workbooks.Open($resolvedOutputPath, 0, $false)
    $templateWorkbook = $excel.Workbooks.Open($resolvedTemplatePath, 0, $true)

    $lookupTytSheet = Get-Worksheet-Safe -Workbook $templateWorkbook -CandidateNames @('REP TYT')
    $lookupPeugSheet = Get-Worksheet-Safe -Workbook $templateWorkbook -CandidateNames @('REP PEUGT')
    $lookupChgnSheet = Get-Worksheet-Safe -Workbook $templateWorkbook -CandidateNames @('REP CHGN')
    $lookupSzkSheet = Get-Worksheet-Safe -Workbook $templateWorkbook -CandidateNames @('REP SZK')

    $repLookups = @{
        tyt = Build-Rep-Lookup -Worksheet $lookupTytSheet
        peug = Build-Rep-Lookup -Worksheet $lookupPeugSheet
        chgn = Build-Rep-Lookup -Worksheet $lookupChgnSheet
        szk = Build-Rep-Lookup -Worksheet $lookupSzkSheet
    }

    $configs = @(
        @{ Key = 'tyt'; Label = 'MATRIZ'; TargetSheet = 'REP TYT'; SourcePath = $resolvedInputTyt },
        @{ Key = 'peug'; Label = 'PEUGEOT'; TargetSheet = 'REP PEUGT'; SourcePath = $resolvedInputPeug },
        @{ Key = 'chgn'; Label = 'CHANGAN'; TargetSheet = 'REP CHGN'; SourcePath = $resolvedInputChgn },
        @{ Key = 'szk'; Label = 'SUZUKI'; TargetSheet = 'REP SZK'; SourcePath = $resolvedInputSzk }
    )

    $repGroupsByKey = @{}

    foreach ($config in $configs) {
        $sheetResult = Process-Rep-Sheet `
            -Excel $excel `
            -OutputWorkbook $outputWorkbook `
            -SourcePath $config.SourcePath `
            -TargetSheetName $config.TargetSheet `
            -Key $config.Key `
            -Label $config.Label `
            -Lookup $repLookups[$config.Key]

        $templateRepSheet = $null
        $outputRepSheet = $null
        try {
            $templateRepSheet = Get-Worksheet-Safe -Workbook $templateWorkbook -CandidateNames @($config.TargetSheet)
            $outputRepSheet = Get-Worksheet-Safe -Workbook $outputWorkbook -CandidateNames @($config.TargetSheet)
            Validate-Rep-Sheet-Output `
                -Worksheet $outputRepSheet `
                -TemplateWorksheet $templateRepSheet `
                -Key $config.Key `
                -Lookup $repLookups[$config.Key] `
                -ExpectedRows $sheetResult.RowCount
        }
        finally {
            if ($null -ne $outputRepSheet) {
                [void][Runtime.InteropServices.Marshal]::ReleaseComObject($outputRepSheet)
            }
            if ($null -ne $templateRepSheet) {
                [void][Runtime.InteropServices.Marshal]::ReleaseComObject($templateRepSheet)
            }
        }

        $repGroupsByKey[$config.Key] = Update-My-Sheet-From-Rep `
            -OutputWorkbook $outputWorkbook `
            -TemplateWorkbook $templateWorkbook `
            -Key $config.Key `
            -RepSheetName $config.TargetSheet

        Validate-My-Sheet-Output `
            -OutputWorkbook $outputWorkbook `
            -TemplateWorkbook $templateWorkbook `
            -Key $config.Key `
            -RepGroups $repGroupsByKey[$config.Key]
    }

    Update-Mayor-Iva-From-Rep `
        -OutputWorkbook $outputWorkbook `
        -TemplateWorkbook $templateWorkbook `
        -RepGroupsByKey $repGroupsByKey

    Validate-Mayor-Iva-Output `
        -OutputWorkbook $outputWorkbook `
        -TemplateWorkbook $templateWorkbook `
        -RepGroupsByKey $repGroupsByKey

    foreach ($sheetName in @('NC REP TYT', 'NC REP PEUG', 'NC REP SZK')) {
        Clear-Nc-Sheet-NoSource -Workbook $outputWorkbook -SheetName $sheetName
    }

    try {
        $outputWorkbook.Application.CalculateFullRebuild()
    }
    catch {
        $outputWorkbook.Application.CalculateFull()
    }

    $outputWorkbook.Save()
}
finally {
    if ($null -ne $lookupSzkSheet) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($lookupSzkSheet)
    }
    if ($null -ne $lookupChgnSheet) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($lookupChgnSheet)
    }
    if ($null -ne $lookupPeugSheet) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($lookupPeugSheet)
    }
    if ($null -ne $lookupTytSheet) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($lookupTytSheet)
    }
    if ($null -ne $templateWorkbook) {
        $templateWorkbook.Close($false)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($templateWorkbook)
    }
    if ($null -ne $outputWorkbook) {
        $outputWorkbook.Close($true)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($outputWorkbook)
    }

    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Output ("OUTPUT|{0}|FACTURACION REPUESTOS TYTSERV" -f (Split-Path -Leaf $resolvedOutputPath))
