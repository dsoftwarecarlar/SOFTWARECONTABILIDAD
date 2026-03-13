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
    $totalRow = Find-Row-Containing -Worksheet $Worksheet -Column 12 -Needle 'TOTAL GENERAL' -StartRow 1 -EndRow ([Math]::Max(200, $lastRow + 10))
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

        if ($refRuc -ne '' -and (Test-Looks-MaskedRuc -Text $currentRuc)) {
            $Worksheet.Cells.Item($row, 9).Value2 = "'" + $refRuc
        }

        if ($refCode -ne '' -and $currentCode -eq '') {
            $Worksheet.Cells.Item($row, 10).Value2 = "'" + $refCode
        }

        if ($refName -ne '' -and (Test-Looks-UnreadableName -Text $currentName)) {
            $Worksheet.Cells.Item($row, 11).Value2 = $refName
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

    switch ($Key) {
        'tyt' {
            $Worksheet.Cells.Item($mayorRow, 16).Formula = "=+'MY REP TYT'!I38"
            $Worksheet.Cells.Item($mayorRow, 18).Formula = "=+'MY REP TYT'!H53"
        }
        'peug' {
            $Worksheet.Cells.Item($mayorRow, 16).Formula = "=+'MY REP PEUG'!I19"
            $Worksheet.Cells.Item($mayorRow, 18).Formula = "=+'MY REP PEUG'!H25"
            $Worksheet.Cells.Item($mayorRow, 25).Formula = "=+X$TotalRow-'NC REP PEUG'!AD9:AE9"
            $Worksheet.Cells.Item($mayorRow, 26).Formula = "=X$TotalRow"
        }
        'chgn' {
            $Worksheet.Cells.Item($mayorRow, 16).Formula = "=+'MY REP CHGN'!J8"
            $Worksheet.Cells.Item($mayorRow, 19).Formula = "=+'MY REP CHGN'!I15"
        }
        'szk' {
            $Worksheet.Cells.Item($mayorRow, 16).Formula = "=+'MY REP SZK'!I31"
            $Worksheet.Cells.Item($mayorRow, 18).Formula = "=+'MY REP SZK'!H50"
        }
        default {
            throw "Clave no soportada: $Key"
        }
    }

    return $mayorRow
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
            $sourceSheet = Get-Worksheet-Safe -Workbook $sourceWorkbook -CandidateNames @('RepLibroVentasGeneral')
        }
        catch {
            $sourceSheet = $sourceWorkbook.Worksheets.Item(1)
        }

        $targetSheet = Get-Worksheet-Safe -Workbook $OutputWorkbook -CandidateNames @($TargetSheetName)
        $oldUsedLastRow = Get-Used-LastRow -Worksheet $targetSheet
        $oldMayorRow = Find-Row-Containing -Worksheet $targetSheet -Column 14 -Needle 'MAYOR' -StartRow 1 -EndRow ([Math]::Max(200, $oldUsedLastRow + 20))

        $sourceLastRow = Get-Used-LastRow -Worksheet $sourceSheet
        Copy-Rep-Range -SourceSheet $sourceSheet -TargetSheet $targetSheet -LastRow $sourceLastRow

        $scanEnd = [Math]::Max($sourceLastRow + 10, $oldUsedLastRow + 10)
        $newTotalRow = Find-Row-Containing -Worksheet $targetSheet -Column 12 -Needle 'TOTAL GENERAL' -StartRow 1 -EndRow $scanEnd
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

    foreach ($config in $configs) {
        Process-Rep-Sheet `
            -Excel $excel `
            -OutputWorkbook $outputWorkbook `
            -SourcePath $config.SourcePath `
            -TargetSheetName $config.TargetSheet `
            -Key $config.Key `
            -Label $config.Label `
            -Lookup $repLookups[$config.Key]
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
