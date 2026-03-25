[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputDir,

    [Parameter(Mandatory = $true)]
    [string]$TemplateDir,

    [Parameter(Mandatory = $false)]
    [string]$RunStamp = '',

    [Parameter(Mandatory = $false)]
    [string]$CancelPath = '',

    [Parameter(Mandatory = $false)]
    [string]$BrandKey = '',

    [Parameter(Mandatory = $false)]
    [string]$FacturaChanganPath = '',

    [Parameter(Mandatory = $false)]
    [string]$NotaChanganPath = '',

    [Parameter(Mandatory = $false)]
    [string]$PxPath = '',

    [Parameter(Mandatory = $false)]
    [string]$RepVtasPath = '',

    [Parameter(Mandatory = $false)]
    [string]$MayorChanganPath = '',

    [Parameter(Mandatory = $false)]
    [string]$FacturaPeugPath = '',

    [Parameter(Mandatory = $false)]
    [string]$NotaPeugPath = '',

    [Parameter(Mandatory = $false)]
    [string]$MayorPeugPath = '',

    [Parameter(Mandatory = $false)]
    [string]$FacturaSzkPath = '',

    [Parameter(Mandatory = $false)]
    [string]$NotaSzkPath = '',

    [Parameter(Mandatory = $false)]
    [string]$MayorSzkPath = '',

    [Parameter(Mandatory = $false)]
    [string]$FacturaTytPath = '',

    [Parameter(Mandatory = $false)]
    [string]$NotaTytPath = '',

    [Parameter(Mandatory = $false)]
    [string]$MayorTytPath = ''
)

$ErrorActionPreference = 'Stop'
$script:CancelFilePath = $CancelPath
$script:StrictValidationEnabled = ([string]$env:SERVICIOS_MARCAS_STRICT_VALIDATE -eq '1')
$script:AllowTemplateDataFallback = ([string]$env:SERVICIOS_MARCAS_ALLOW_TEMPLATE_FALLBACK -eq '1')
$script:PivotRefreshEnabled = ([string]$env:SERVICIOS_MARCAS_REFRESH_PIVOTS -eq '1')
$script:LayoutRestoreEnabled = ([string]$env:SERVICIOS_MARCAS_RESTORE_LAYOUT -eq '1')

function Test-CancelRequested {
    if ([string]::IsNullOrWhiteSpace($script:CancelFilePath)) {
        return $false
    }

    return (Test-Path -LiteralPath $script:CancelFilePath)
}

function Assert-NotCancelled {
    param([string]$Context = 'proceso')

    if (Test-CancelRequested) {
        throw "CANCELLED: Proceso detenido por el usuario."
    }
}

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

function Get-TemplateKey-FromAgency {
    param([string]$Agency)

    $agencyText = (Normalize-Text $Agency).ToUpper()
    switch ($agencyText) {
        'CHANGAN' { return 'changan' }
        'PEUGEOT' { return 'peug' }
        'MATRIZ' { return 'tyt' }
        'SUZUKI AMBATO' { return 'szk' }
        'SUZUKI RIOBAMBA' { return 'szk' }
        default { return '' }
    }
}

function Parse-TabFile {
    param(
        [string]$Path,
        [string]$Kind = ''
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return @()
    }

    if (-not (Test-Path -LiteralPath $Path)) {
        return @()
    }

    try {
        $rows = Import-Csv -LiteralPath $Path -Delimiter "`t" -Encoding UTF8
        # Filtrar filas sin agencia ni documento.
        $filtered = @($rows | Where-Object { (Normalize-Text $_.'Agencia') -ne '' -or (Normalize-Text $_.'Agencia :') -ne '' -or (Normalize-Text $_.'Factura') -ne '' -or (Normalize-Text $_.'Nota Cred.') -ne '' })
        if ($filtered.Count -gt 0) {
            return $filtered
        }
    } catch {
    }

    $fallbackRows = New-Object System.Collections.Generic.List[object]
    foreach ($line in Get-Content -LiteralPath $Path -Encoding UTF8) {
        $columns = @($line -split "`t")
        if ($columns.Count -eq 0) {
            continue
        }

        if ($Kind -eq 'factura') {
            if ($columns.Count -lt 36) {
                continue
            }

            if ((Normalize-Text $columns[0]).ToUpperInvariant() -ne 'AGENCIA :' -or (Normalize-Text $columns[2]).ToUpperInvariant() -ne 'SERIE') {
                continue
            }

            $fallbackRows.Add([pscustomobject]@{
                'Agencia :' = Normalize-Text $columns[1]
                'Serie' = Normalize-Text $columns[19]
                'Factura' = Normalize-Text $columns[20]
                'Fecha' = Normalize-Text $columns[21]
                'Orden' = Normalize-Text $columns[22]
                'C.I.' = Normalize-Text $columns[23]
                'Cliente' = Normalize-Text $columns[24]
                'Sub total' = Normalize-Text $columns[25]
                'Des-cuento' = Normalize-Text $columns[26]
                'Iva 12%' = Normalize-Text $columns[29]
                'Iva 15%' = Normalize-Text $columns[30]
                'Total' = Normalize-Text $columns[32]
                'Asiento' = Normalize-Text $columns[33]
                'Gar.Ext.' = Normalize-Text $columns[34]
                'T.V.' = Normalize-Text $columns[35]
            }) | Out-Null
            continue
        }

        if ($Kind -eq 'nota') {
            if ($columns.Count -lt 40) {
                continue
            }

            if ((Normalize-Text $columns[0]).ToUpperInvariant() -ne 'AGENCIA :' -or (Normalize-Text $columns[2]).ToUpperInvariant() -ne 'NOTA CRED.') {
                continue
            }

            $fallbackRows.Add([pscustomobject]@{
                'Agencia :' = Normalize-Text $columns[1]
                'Nota Cred.' = Normalize-Text $columns[20]
                'Fecha' = Normalize-Text $columns[21]
                'Tipo' = Normalize-Text $columns[22]
                'Serie' = Normalize-Text $columns[23]
                'Factura' = Normalize-Text $columns[24]
                'Orden' = Normalize-Text $columns[25]
                'Cedula' = Normalize-Text $columns[26]
                'Cliente' = Normalize-Text $columns[27]
                'Sub total' = Normalize-Text $columns[28]
                'Des-cuento' = Normalize-Text $columns[29]
                'Iva 15%' = Normalize-Text $columns[32]
                'Iva 12 %' = Normalize-Text $columns[33]
                'Total' = Normalize-Text $columns[35]
                'Anticipo' = Normalize-Text $columns[36]
                'NETO' = Normalize-Text $columns[37]
                'Asiento' = Normalize-Text $columns[38]
                'Gar.Ext.' = Normalize-Text $columns[39]
            }) | Out-Null
        }
    }

    return $fallbackRows.ToArray()
}

function Build-SourceRows-FromBrandInputs {
    param(
        [hashtable]$BrandFileMap,
        [string]$BrandKey = ''
    )

    $rows = New-Object System.Collections.Generic.List[object]
    $brandOrder = @('changan', 'peug', 'szk', 'tyt')
    if (-not [string]::IsNullOrWhiteSpace($BrandKey)) {
        $brandOrder = @($brandOrder | Where-Object { $_ -eq $BrandKey })
    }

    $appendRows = {
        param(
            [object[]]$Items,
            [string]$DocType,
            [string]$TemplateKey
        )

        $index = $rows.Count + 1
        foreach ($item in $Items) {
            $agencyRaw = $item.'Agencia :'
            if (-not $agencyRaw) { $agencyRaw = $item.'Agencia' }
            $agency = Normalize-Text $agencyRaw
            if ($agency -eq '' -or [string]::IsNullOrWhiteSpace($TemplateKey)) {
                continue
            }

            $seriesRaw = Normalize-Text ($item.'Serie')
            $documentRaw = if ($DocType -eq 'DC') { $item.'Nota Cred.' } else { $item.'Factura' }
            if (-not $documentRaw) { $documentRaw = $item.'Factura' }
            if (-not $documentRaw) { $documentRaw = $item.'Nota Cred.' }
            $document = Normalize-Text $documentRaw
            if ($document -eq '') { continue }

            $orderRaw = Normalize-Text ($item.'Orden')
            $cedulaRaw = $item.'Cedula'
            if (-not $cedulaRaw) { $cedulaRaw = $item.'C.I.' }
            $cedula = Normalize-Text $cedulaRaw
            $customerRaw = Normalize-Text ($item.'Cliente')
            $affectedDocumentRaw = ''
            if ($DocType -eq 'DC') {
                $affectedDocumentRaw = Normalize-Text ($item.'Factura')
            }

            $fechaText = Normalize-Text ($item.'Fecha')
            $dateFactValue = $null
            $dateNoteValue = $null
            if ($fechaText -ne '') {
                $parsed = Get-Date -Date $fechaText -ErrorAction SilentlyContinue
                if ($parsed) {
                    if ($DocType -eq 'DC') {
                        $dateNoteValue = $parsed.ToOADate()
                    } else {
                        $dateFactValue = $parsed.ToOADate()
                    }
                }
            }

            $subtotal = [double]([decimal]($item.'Sub total'))
            $discount = [double]([decimal]($item.'Des-cuento'))
            $iva12 = [double]([decimal]($item.'Iva 12%'))
            if (-not $iva12) { $iva12 = [double]([decimal]($item.'Iva 12 %')) }
            $iva15 = [double]([decimal]($item.'Iva 15%'))
            $iva = [double]([decimal]($iva12 + $iva15))
            $interest = [double]([decimal]($item.'Interes'))
            $total = [double]([decimal]($item.'Total'))
            $garExt = Normalize-Text ($item.'Gar.Ext.')
            $tv = Normalize-Text ($item.'T.V.')
            $anticipo = [double]([decimal]($item.'Anticipo'))
            $neto = [double]([decimal]($item.'NETO'))
            $netoConIva = [double]([decimal]($item.'Netocon Iva'))
            $netoSinIva = [double]([decimal]($item.'Netosin Iva'))
            $netoIva0 = [double]([decimal]($item.'Neto Iva 0'))

            $noteCredit = $discount
            $manoObra = $subtotal
            if (-not $manoObra) {
                if ($DocType -eq 'DC') {
                    $manoObra = $netoConIva + $netoSinIva + $discount
                } else {
                    $manoObra = $netoConIva + $netoIva0 + $discount
                }
            }

            if (-not $subtotal) { $subtotal = 0.0 }
            if (-not $noteCredit) { $noteCredit = 0.0 }
            if (-not $iva12) { $iva12 = 0.0 }
            if (-not $iva15) { $iva15 = 0.0 }
            if (-not $iva) { $iva = 0.0 }
            if (-not $interest) { $interest = 0.0 }
            if (-not $total) { $total = 0.0 }
            if (-not $anticipo) { $anticipo = 0.0 }
            if (-not $neto) { $neto = 0.0 }
            if (-not $manoObra) { $manoObra = $subtotal }

            $rows.Add([pscustomobject]@{
                RowIndex = $index
                TemplateKey = $TemplateKey
                Agency = $agency
                AgencyRaw = $agencyRaw
                Center = ''
                CenterRaw = ''
                Order = $orderRaw
                OrderRaw = $orderRaw
                Advisor = ''
                AdvisorRaw = ''
                Line = ''
                LineRaw = ''
                DocType = $DocType
                Cedula = $cedula
                CedulaRaw = $cedulaRaw
                Customer = $customerRaw
                CustomerRaw = $customerRaw
                DocumentRaw = $document
                DocumentTrim = (Trim-Document $document)
                Series = $seriesRaw
                SeriesRaw = $seriesRaw
                FormaPago = $tv
                Authorization = ''
                DateFactValue = $dateFactValue
                DateNoteValue = $dateNoteValue
                NoteCredit = $noteCredit
                TotalManoObra = $manoObra
                TotalSubcontratos = 0.0
                TotalInsumos = 0.0
                TotalServicio = 0.0
                TotalAccesorios = 0.0
                TotalRepuestos = 0.0
                Interes = $interest
                Iva = $iva
                Iva12 = $iva12
                Iva15 = $iva15
                Total = $total
                Anticipo = $anticipo
                Neto = $neto
                Costo = 0.0
                CostoLubricantes = 0.0
                CostoAccesorios = 0.0
                CostoRepuestos = 0.0
                CostoPintura = 0.0
                CostoSubconNc = 0.0
                GarExt = $garExt
                GarExtRaw = $garExt
                AffectedDocumentTrim = (Trim-Document $affectedDocumentRaw)
                AffectedDocumentRaw = $affectedDocumentRaw
                MotivoNc = ''
                ObservacionNc = ''
            }) | Out-Null
            $index++
        }
    }

    foreach ($targetBrand in $brandOrder) {
        if (-not $BrandFileMap.ContainsKey($targetBrand)) {
            continue
        }

        $paths = $BrandFileMap[$targetBrand]
        $facturas = Parse-TabFile -Path (Normalize-Text $paths.FacturaPath) -Kind 'factura'
        $notas = Parse-TabFile -Path (Normalize-Text $paths.NotaPath) -Kind 'nota'

        & $appendRows $facturas 'FA' $targetBrand
        & $appendRows $notas 'DC' $targetBrand
    }

    return $rows.ToArray()
}
function Copy-File-WithRetry {
    param(
        [string]$SourcePath,
        [string]$DestinationPath,
        [int]$Attempts = 8
    )

    $lastError = ''
    for ($attempt = 1; $attempt -le $Attempts; $attempt++) {
        try {
            Copy-Item -LiteralPath $SourcePath -Destination $DestinationPath -Force -ErrorAction Stop
            return
        } catch {
            $lastError = $_.Exception.Message
            if ($attempt -lt $Attempts) {
                Start-Sleep -Milliseconds (700 * $attempt)
            }
        }
    }

    throw ("No se pudo preparar copia temporal del archivo fuente. Origen: {0}. Destino: {1}. Detalle: {2}" -f $SourcePath, $DestinationPath, $lastError)
}

function Resolve-PythonBinary {
    $candidates = New-Object System.Collections.Generic.List[string]

    try {
        $pythonCommand = Get-Command python.exe -ErrorAction Stop
        if ($null -ne $pythonCommand -and -not [string]::IsNullOrWhiteSpace($pythonCommand.Source)) {
            $candidates.Add($pythonCommand.Source) | Out-Null
        }
    } catch {
    }

    try {
        $pythonCommand = Get-Command python -ErrorAction Stop
        if ($null -ne $pythonCommand -and -not [string]::IsNullOrWhiteSpace($pythonCommand.Source)) {
            $candidates.Add($pythonCommand.Source) | Out-Null
        }
    } catch {
    }

    if (-not [string]::IsNullOrWhiteSpace($env:PYTHON_BINARY) -and (Test-Path -LiteralPath $env:PYTHON_BINARY)) {
        $candidates.Add($env:PYTHON_BINARY) | Out-Null
    }

    $programFiles = if ([string]::IsNullOrWhiteSpace($env:ProgramFiles)) { 'C:\Program Files' } else { $env:ProgramFiles }
    $programFilesX86 = if ([string]::IsNullOrWhiteSpace(${env:ProgramFiles(x86)})) { 'C:\Program Files (x86)' } else { ${env:ProgramFiles(x86)} }

    foreach ($candidate in @(
        (Join-Path $programFiles 'Python312\python.exe'),
        (Join-Path $programFiles 'Python311\python.exe'),
        (Join-Path $programFiles 'Python310\python.exe'),
        (Join-Path $programFilesX86 'Python312\python.exe'),
        (Join-Path $programFilesX86 'Python311\python.exe'),
        (Join-Path $programFilesX86 'Python310\python.exe')
    )) {
        if (-not [string]::IsNullOrWhiteSpace($candidate) -and (Test-Path -LiteralPath $candidate)) {
            $candidates.Add($candidate) | Out-Null
        }
    }

    foreach ($candidate in $candidates) {
        if (-not [string]::IsNullOrWhiteSpace($candidate) -and (Test-Path -LiteralPath $candidate)) {
            return $candidate
        }
    }

    throw 'No se encontro Python para leer el Excel fuente sin COM.'
}

function Read-SourceRows-FromFile {
    param(
        [string]$InputPath,
        [string]$WorkingDirectory
    )

    $pythonBinary = Resolve-PythonBinary
    $readerScript = Resolve-RequiredPath -Path (Join-Path $PSScriptRoot 'python_services\processors\servicios_marcas\readers.py') -Label 'readers.py'
    $jsonPath = Join-Path $WorkingDirectory 'source_rows.json'
    if (Test-Path -LiteralPath $jsonPath) {
        Remove-Item -LiteralPath $jsonPath -Force -ErrorAction SilentlyContinue
    }

    $pythonOutput = & $pythonBinary $readerScript 'source' '--input' $InputPath '--output-json' $jsonPath 2>&1
    if ($LASTEXITCODE -ne 0) {
        $detail = (($pythonOutput | ForEach-Object { "$_" }) -join [Environment]::NewLine).Trim()
        if ($detail -eq '') {
            $detail = 'Python termino sin detalle.'
        }
        throw ("No se pudo leer el Excel fuente sin COM. Detalle: {0}" -f $detail)
    }

    if (-not (Test-Path -LiteralPath $jsonPath)) {
        throw 'El lector del Excel fuente no genero source_rows.json.'
    }

    $payloadText = Get-Content -LiteralPath $jsonPath -Raw -Encoding UTF8
    if ([string]::IsNullOrWhiteSpace($payloadText)) {
        throw 'El lector del Excel fuente devolvio un JSON vacio.'
    }

    $payload = $payloadText | ConvertFrom-Json
    $rows = @()
    if ($null -ne $payload -and $null -ne $payload.rows) {
        $rows = @($payload.rows)
    }

    Write-Output ("INFO|source_read|rows={0}" -f $rows.Count)
    return $rows
}

function Read-PxRows {
    param(
        [string]$PxPath,
        [string]$BrandKey,
        [string]$WorkingDirectory
    )

    if ([string]::IsNullOrWhiteSpace($PxPath) -or -not (Test-Path -LiteralPath $PxPath)) {
        return @()
    }

    $pythonBinary = Resolve-PythonBinary
    $readerScript = Resolve-RequiredPath -Path (Join-Path $PSScriptRoot 'python_services\processors\servicios_marcas\readers.py') -Label 'readers.py'
    $jsonPath = Join-Path $WorkingDirectory 'px_rows.json'
    if (Test-Path -LiteralPath $jsonPath) {
        Remove-Item -LiteralPath $jsonPath -Force -ErrorAction SilentlyContinue
    }

    $pythonOutput = & $pythonBinary $readerScript 'px' '--input' $PxPath '--brand' $BrandKey '--output-json' $jsonPath 2>&1
    if ($LASTEXITCODE -ne 0) {
        $detail = (($pythonOutput | ForEach-Object { "$_" }) -join [Environment]::NewLine).Trim()
        if ($detail -eq '') { $detail = 'Python termino sin detalle.' }
        throw ("No se pudo leer el archivo PX. Detalle: {0}" -f $detail)
    }

    if (-not (Test-Path -LiteralPath $jsonPath)) {
        return @()
    }

    $payloadText = Get-Content -LiteralPath $jsonPath -Raw -Encoding UTF8
    if ([string]::IsNullOrWhiteSpace($payloadText)) {
        return @()
    }

    $payload = $payloadText | ConvertFrom-Json
    if ($null -eq $payload -or $null -eq $payload.rows) {
        return @()
    }

    return @($payload.rows)
}

function Get-BrandDisplayLabel {
    param([string]$BrandKey)

    switch ((Normalize-Text $BrandKey).ToLowerInvariant()) {
        'changan' { return 'CHANGAN' }
        'peug' { return 'PEUGEOT' }
        'szk' { return 'SUZUKI' }
        'tyt' { return 'TOYOTA' }
        default { return (Normalize-Text $BrandKey).ToUpperInvariant() }
    }
}

function Get-IndexedRowValue {
    param(
        [object]$Row,
        [int]$Index
    )

    if ($null -eq $Row) {
        return ''
    }

    if ($Row -is [System.Collections.IList] -and $Index -lt $Row.Count) {
        return Normalize-Text $Row[$Index]
    }

    return ''
}

function Convert-PxRowsToDetailRows {
    param([object[]]$Rows)

    $detailRows = New-Object System.Collections.Generic.List[object]
    foreach ($row in @($Rows)) {
        $ag = Get-IndexedRowValue -Row $row -Index 1
        $factura = Get-IndexedRowValue -Row $row -Index 4
        $item = Get-IndexedRowValue -Row $row -Index 15

        if ($ag -eq '' -or $factura -eq '' -or $item -eq '') {
            continue
        }

        if ($ag -notmatch '^\d+$') {
            continue
        }

        $detailRows.Add([pscustomobject]@{
            Agencia = $ag
            Estado = Get-IndexedRowValue -Row $row -Index 2
            Orden = Get-IndexedRowValue -Row $row -Index 3
            Factura = $factura
            Fecha = Get-IndexedRowValue -Row $row -Index 6
            PxNo = Get-IndexedRowValue -Row $row -Index 8
            Cuenta = Get-IndexedRowValue -Row $row -Index 9
            Fr = Get-IndexedRowValue -Row $row -Index 11
            Codigo = Get-IndexedRowValue -Row $row -Index 13
            Item = $item
            PvpBruto = To-Number (Get-IndexedRowValue -Row $row -Index 19)
            DescPct = To-Number (Get-IndexedRowValue -Row $row -Index 21)
            DescValor = To-Number (Get-IndexedRowValue -Row $row -Index 22)
            PvpNeto = To-Number (Get-IndexedRowValue -Row $row -Index 23)
            Costo = To-Number (Get-IndexedRowValue -Row $row -Index 24)
            Origen = Get-IndexedRowValue -Row $row -Index 25
        }) | Out-Null
    }

    return $detailRows.ToArray()
}

function Get-PxDetailRanges {
    param([object]$Worksheet)

    $ranges = New-Object System.Collections.Generic.List[object]
    $seen = @{}
    $usedRange = $null
    try {
        $usedRange = $Worksheet.UsedRange
        $lastRow = [int]($usedRange.Row + $usedRange.Rows.Count - 1)
        for ($row = 1; $row -le $lastRow; $row++) {
            $formulaText = Normalize-Text $Worksheet.Cells.Item($row, 12).Formula
            if ($formulaText -eq '') {
                continue
            }

            $normalizedFormula = $formulaText.ToUpperInvariant().Replace('$', '')
            if ($normalizedFormula -notmatch 'SUBTOTAL\(9,L(\d+):L(\d+)\)') {
                continue
            }

            $startRow = [int]$Matches[1]
            $endRow = [int]$Matches[2]
            $key = '{0}:{1}' -f $startRow, $endRow
            if ($seen.ContainsKey($key)) {
                continue
            }

            $seen[$key] = $true
            $ranges.Add([pscustomobject]@{
                StartRow = $startRow
                EndRow = $endRow
            }) | Out-Null
        }
    } finally {
        if ($null -ne $usedRange) {
            [void][Runtime.Interopservices.Marshal]::ReleaseComObject($usedRange)
        }
    }

    return @($ranges.ToArray() | Sort-Object @{ Expression = { [int]$_.StartRow } })
}

function Write-PxRows-ToWorksheet {
    param(
        [object]$Worksheet,
        [object[]]$Rows,
        [string]$BrandKey = ''
    )

    $ranges = @(Get-PxDetailRanges -Worksheet $Worksheet)
    if ($ranges.Count -lt 2) {
        throw "La hoja PX no contiene los bloques esperados para notas de credito y ventas por liquidar."
    }

    $topRange = $ranges[0]
    $bottomRange = $ranges[1]

    foreach ($pxClearRange in @($topRange, $bottomRange)) {
        $pxRangeRef = $Worksheet.Range("B$($pxClearRange.StartRow):Q$($pxClearRange.EndRow)")
        try {
            $pxConstants = $pxRangeRef.SpecialCells(2)
            $null = $pxConstants.ClearContents()
            [void][Runtime.Interopservices.Marshal]::ReleaseComObject($pxConstants)
        } catch {}
        [void][Runtime.Interopservices.Marshal]::ReleaseComObject($pxRangeRef)
    }

    $detailRows = @(Convert-PxRowsToDetailRows -Rows $Rows)
    $capacity = [int]($bottomRange.EndRow - $bottomRange.StartRow + 1)
    if ($detailRows.Count -gt $capacity) {
        throw ("Las filas PX de {0} exceden la capacidad de la plantilla. Capacidad={1}, filas={2}" -f $BrandKey, $capacity, $detailRows.Count)
    }

    if ($detailRows.Count -gt 0) {
        $sheetRows = New-Object System.Collections.Generic.List[object]
        foreach ($row in $detailRows) {
            $sheetRows.Add(@(
                (Get-Excel-TextLiteral $row.Agencia),
                (Get-Excel-TextLiteral $row.Estado),
                (Get-Excel-TextLiteral $row.Orden),
                (Get-Excel-TextLiteral $row.Factura),
                (Get-Excel-TextLiteral $row.Fecha),
                (Get-Excel-TextLiteral $row.PxNo),
                (Get-Excel-TextLiteral $row.Cuenta),
                (Get-Excel-TextLiteral $row.Fr),
                (Get-Excel-TextLiteral $row.Codigo),
                (Get-Excel-TextLiteral $row.Item),
                (Get-NumericMatrixValue -Value $row.PvpBruto -BlankIfZero),
                (Get-NumericMatrixValue -Value $row.DescPct -BlankIfZero),
                (Get-NumericMatrixValue -Value $row.DescValor -BlankIfZero),
                (Get-NumericMatrixValue -Value $row.PvpNeto -BlankIfZero),
                (Get-NumericMatrixValue -Value $row.Costo -BlankIfZero),
                (Get-Excel-TextLiteral $row.Origen)
            )) | Out-Null
        }

        Write-Rows-ToWorksheet -Worksheet $Worksheet -Rows $sheetRows.ToArray() -StartRow ([int]$bottomRange.StartRow) -StartColumn 2
    }

    $processCell = $Worksheet.Cells.Item(3, 4)
    try {
        $null = $processCell.Value2 = (Get-Excel-TextLiteral ((Get-Date).ToString('dd/MM/yyyy HH.mm.ss')))
    } finally {
        [void][Runtime.Interopservices.Marshal]::ReleaseComObject($processCell)
    }

    $null = $Worksheet.Cells.Item(5, 4).Value2 = (Get-Excel-TextLiteral (Get-BrandDisplayLabel -BrandKey $BrandKey))

    return [pscustomobject]@{
        RowCount = $detailRows.Count
        TopCapacity = [int]($topRange.EndRow - $topRange.StartRow + 1)
        BottomCapacity = $capacity
    }
}

function Update-PxPeriodAnchor {
    param(
        [object]$Worksheet,
        [object]$PeriodDateValue
    )

    if ($null -eq $Worksheet -or $null -eq $PeriodDateValue -or $PeriodDateValue -eq '') {
        return
    }

    $periodDate = [datetime]::FromOADate([double]$PeriodDateValue)
    $periodStart = Get-Date -Year $periodDate.Year -Month $periodDate.Month -Day 1
    Set-DateCellValue -Worksheet $Worksheet -Row 3 -Column 8 -Value $periodStart.ToOADate() -Context 'PX_PERIODO_INICIO'
}

function New-GetPivotDataFormula {
    param(
        [string]$DataField,
        [string]$Anchor,
        [string]$Account,
        [string]$Description,
        [string]$Sign = '+'
    )

    $normalizedSign = if ($Sign -eq '-') { '-' } else { '+' }
    $safeDataField = ([string]$DataField).Replace('"', '""')
    $safeAccount = ([string]$Account).Replace('"', '""')
    $safeDescription = ([string]$Description).Replace('"', '""')

    return ('{0}IFERROR(GETPIVOTDATA("{1}",{2},"CUENTA","{3}","DESCRIPCION","{4}"),0)' -f $normalizedSign, $safeDataField, $Anchor, $safeAccount, $safeDescription)
}

function New-SumIfAccountFormula {
    param(
        [string]$SheetName,
        [string]$SumColumn,
        [string]$Account,
        [string]$Sign = '+'
    )

    $normalizedSign = if ($Sign -eq '-') { '-' } else { '+' }
    $safeSheetName = ([string]$SheetName).Replace("'", "''")
    $safeAccount = ([string]$Account).Replace('"', '""')
    return ('{0}SUMIFS(''{1}''!${2}:${2},''{1}''!$E:$E,"{3}")' -f $normalizedSign, $safeSheetName, $SumColumn, $safeAccount)
}

function Get-FirstGeneratedAccount {
    param(
        [object[]]$Rows,
        [string]$Pattern
    )

    $match = @(
        @($Rows) |
            Where-Object { (Normalize-Text $_.Account) -match $Pattern } |
            Select-Object -First 1
    )[0]

    if ($null -eq $match) {
        return ''
    }

    return Normalize-Text $match.Account
}

function Update-PrecontVentasControlFormulas {
    param(
        [object]$Worksheet,
        [object[]]$Rows
    )

    if ($null -eq $Worksheet) {
        return
    }

    $rowList = @($Rows)
    $creditClientAccount = Get-FirstGeneratedAccount -Rows $rowList -Pattern '^010104\d{2}0002$'
    $reserveAccount = Get-FirstGeneratedAccount -Rows $rowList -Pattern '^020105\d{2}000\d$'
    $liquidarCreditAccount = Get-FirstGeneratedAccount -Rows $rowList -Pattern '^020120\d{2}0002$'
    $liquidarDiscountAccount = Get-FirstGeneratedAccount -Rows $rowList -Pattern '^020120\d{2}0004$'

    if ($creditClientAccount -ne '' -and $reserveAccount -ne '') {
        $formula = '=' +
            (New-SumIfAccountFormula -SheetName 'PrecontabilizacionVentas' -SumColumn 'I' -Account $creditClientAccount -Sign '+') +
            (New-SumIfAccountFormula -SheetName 'PrecontabilizacionVentas' -SumColumn 'H' -Account $reserveAccount -Sign '-')
        $Worksheet.Range('V7').Formula = $formula
    } else {
        $Worksheet.Range('V7').Formula = '=0'
    }

    if ($liquidarCreditAccount -ne '' -and $liquidarDiscountAccount -ne '') {
        $formula = '=' +
            (New-SumIfAccountFormula -SheetName 'PrecontabilizacionVentas' -SumColumn 'I' -Account $liquidarCreditAccount -Sign '+') +
            (New-SumIfAccountFormula -SheetName 'PrecontabilizacionVentas' -SumColumn 'H' -Account $liquidarDiscountAccount -Sign '-')
        $Worksheet.Range('V10').Formula = $formula
    } else {
        $Worksheet.Range('V10').Formula = '=0'
    }
}

function Update-RepAndNoteControlFormulas {
    param(
        [object]$RepWorksheet,
        [object]$NoteWorksheet,
        [object[]]$Rows
    )

    $rowList = @($Rows)
    $salesContadoAccount = Get-FirstGeneratedAccount -Rows $rowList -Pattern '^040101\d{2}0001$'
    $salesCreditoAccount = Get-FirstGeneratedAccount -Rows $rowList -Pattern '^040101\d{2}0003$'
    $discountContadoAccount = Get-FirstGeneratedAccount -Rows $rowList -Pattern '^040101\d{2}0010$'
    $discountCreditoAccount = Get-FirstGeneratedAccount -Rows $rowList -Pattern '^040101\d{2}0012$'
    $returnContadoAccount = Get-FirstGeneratedAccount -Rows $rowList -Pattern '^040101\d{2}0014$'

    if ($null -ne $RepWorksheet) {
        $repSalesFormula = '=0'
        $repDiscountFormula = '=0'
        if ($salesContadoAccount -ne '' -or $salesCreditoAccount -ne '') {
            $repSalesFormula = '='
            if ($salesContadoAccount -ne '') {
                $repSalesFormula += (New-SumIfAccountFormula -SheetName 'PrecontabilizacionVentas' -SumColumn 'I' -Account $salesContadoAccount -Sign '+')
            }
            if ($salesCreditoAccount -ne '') {
                $repSalesFormula += (New-SumIfAccountFormula -SheetName 'PrecontabilizacionVentas' -SumColumn 'I' -Account $salesCreditoAccount -Sign '+')
            }
        }
        if ($discountContadoAccount -ne '' -or $discountCreditoAccount -ne '') {
            $repDiscountFormula = '='
            if ($discountContadoAccount -ne '') {
                $repDiscountFormula += (New-SumIfAccountFormula -SheetName 'PrecontabilizacionVentas' -SumColumn 'H' -Account $discountContadoAccount -Sign '+')
            }
            if ($discountCreditoAccount -ne '') {
                $repDiscountFormula += (New-SumIfAccountFormula -SheetName 'PrecontabilizacionVentas' -SumColumn 'H' -Account $discountCreditoAccount -Sign '+')
            }
        }

        $RepWorksheet.Range('D9').Formula = $repSalesFormula
        $RepWorksheet.Range('E9').Formula = $repDiscountFormula
    }

    if ($null -ne $NoteWorksheet) {
        $noteReturnFormula = '=0'
        $noteDiscountFormula = '=0'
        if ($returnContadoAccount -ne '') {
            $noteReturnFormula = '=' + (New-SumIfAccountFormula -SheetName 'PrecontabilizacionVentas' -SumColumn 'H' -Account $returnContadoAccount -Sign '+')
        }
        if ($discountContadoAccount -ne '' -or $discountCreditoAccount -ne '') {
            $noteDiscountFormula = '='
            if ($discountContadoAccount -ne '') {
                $noteDiscountFormula += (New-SumIfAccountFormula -SheetName 'PrecontabilizacionVentas' -SumColumn 'I' -Account $discountContadoAccount -Sign '+')
            }
            if ($discountCreditoAccount -ne '') {
                $noteDiscountFormula += (New-SumIfAccountFormula -SheetName 'PrecontabilizacionVentas' -SumColumn 'I' -Account $discountCreditoAccount -Sign '+')
            }
        }

        $NoteWorksheet.Range('F4').Formula = $noteReturnFormula
        $NoteWorksheet.Range('G4').Formula = $noteDiscountFormula
    }
}

function Read-MayorRows {
    param(
        [string]$MayorPath,
        [string]$WorkingDirectory
    )

    if ([string]::IsNullOrWhiteSpace($MayorPath) -or -not (Test-Path -LiteralPath $MayorPath)) {
        return @()
    }

    $pythonBinary = Resolve-PythonBinary
    $readerScript = Resolve-RequiredPath -Path (Join-Path $PSScriptRoot 'python_services\processors\servicios_marcas\readers.py') -Label 'readers.py'
    $jsonPath = Join-Path $WorkingDirectory ("mayor_{0}.json" -f ([Guid]::NewGuid().ToString('N')))
    if (Test-Path -LiteralPath $jsonPath) {
        Remove-Item -LiteralPath $jsonPath -Force -ErrorAction SilentlyContinue
    }

    $pythonOutput = & $pythonBinary $readerScript 'mayor' '--input' $MayorPath '--output-json' $jsonPath 2>&1
    if ($LASTEXITCODE -ne 0) {
        $detail = (($pythonOutput | ForEach-Object { "$_" }) -join [Environment]::NewLine).Trim()
        if ($detail -eq '') { $detail = 'Python termino sin detalle.' }
        throw ("No se pudo leer el mayor TXT. Detalle: {0}" -f $detail)
    }

    if (-not (Test-Path -LiteralPath $jsonPath)) {
        return @()
    }

    $payloadText = Get-Content -LiteralPath $jsonPath -Raw -Encoding UTF8
    if ([string]::IsNullOrWhiteSpace($payloadText)) {
        return @()
    }

    $payload = $payloadText | ConvertFrom-Json
    if ($null -eq $payload -or $null -eq $payload.rows) {
        return @()
    }

    return @($payload.rows)
}

function Test-MayorPxAdjustmentRow {
    param([object]$Row)

    if ($null -eq $Row) {
        return $false
    }

    $account = Get-CompactAccountCode $Row.account
    $origin = (Normalize-Text $Row.origin).ToUpperInvariant()
    $seat = Normalize-Text $Row.seat
    $detail = (Normalize-Text $Row.detail).ToUpperInvariant()

    if ($account -notmatch '^040101\d{2}(0003|0012)$') {
        return $false
    }

    if ($detail -match 'REGISTRO DE PX AJUSTE DE EGRESO') {
        return $true
    }

    return ($origin -eq 'AGCM' -and $seat -eq '435')
}

function Filter-MayorRowsForWorkbook {
    param([object[]]$Rows)

    $kept = New-Object System.Collections.Generic.List[object]
    $removed = New-Object System.Collections.Generic.List[object]
    $balanceAdjustments = @{}

    foreach ($row in @($Rows)) {
        $account = Get-CompactAccountCode $row.account
        if (-not $balanceAdjustments.ContainsKey($account)) {
            $balanceAdjustments[$account] = [double]0
        }

        if (Test-MayorPxAdjustmentRow -Row $row) {
            $balanceAdjustments[$account] = [double]$balanceAdjustments[$account] + [double](To-Number $row.credit) - [double](To-Number $row.debit)
            $removed.Add($row) | Out-Null
            continue
        }

        $clone = $row | Select-Object *
        $adjustedBalance = [double](To-Number $row.balance) + [double]$balanceAdjustments[$account]
        try {
            $clone | Add-Member -NotePropertyName effective_balance -NotePropertyValue ([double](Round-Amount $adjustedBalance)) -Force
        } catch {}
        $kept.Add($clone) | Out-Null
    }

    return [pscustomobject]@{
        Rows = $kept.ToArray()
        Removed = $removed.ToArray()
    }
}

function Split-MayorUnmappedAccounts {
    param([string[]]$Accounts)

    $fatal = New-Object System.Collections.Generic.List[string]
    $warning = New-Object System.Collections.Generic.List[string]

    foreach ($account in @($Accounts | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Select-Object -Unique)) {
        $normalized = Normalize-Text $account
        if ($normalized -match '^04\.01\.01\.\d{2}\.0002$') {
            $warning.Add($normalized) | Out-Null
            continue
        }

        $fatal.Add($normalized) | Out-Null
    }

    return [pscustomobject]@{
        Warning = $warning.ToArray()
        Fatal = $fatal.ToArray()
    }
}

function Read-TabRows {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) {
        return @()
    }

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($line in Get-Content -LiteralPath $Path -Encoding UTF8) {
        $rows.Add(($line -split "`t")) | Out-Null
    }

    return $rows.ToArray()
}

function Open-Workbook-WithRetry {
    param(
        [object]$Excel,
        [string]$Path,
        [bool]$ReadOnly = $true,
        [int]$Attempts = 20
    )

    $lastError = ''
    for ($attempt = 1; $attempt -le $Attempts; $attempt++) {
        try {
            return $Excel.Workbooks.Open($Path, 0, $ReadOnly)
        } catch {
            $lastError = $_.Exception.Message
            if ($attempt -lt $Attempts) {
                if ($lastError -match 'RPC_E_CALL_REJECTED|0x80010001|0x800AC472') {
                    Start-Sleep -Milliseconds (900 * $attempt)
                } else {
                    Start-Sleep -Milliseconds (700 * $attempt)
                }
            }
        }
    }

    throw ("No se pudo abrir el libro de Excel en {0}. Cierra archivos abiertos o procesos Excel colgados y reintenta. Detalle: {1}" -f $Path, $lastError)
}

function Save-Workbook-WithRetry {
    param(
        [object]$Workbook,
        [string]$PathForError,
        [int]$Attempts = 6
    )

    $lastError = ''
    for ($attempt = 1; $attempt -le $Attempts; $attempt++) {
        try {
            $Workbook.Save()
            return
        } catch {
            $lastError = $_.Exception.Message
            if ($attempt -lt $Attempts) {
                Start-Sleep -Milliseconds (600 * $attempt)
            }
        }
    }

    throw ("No se pudo guardar el archivo de salida en {0}. Detalle: {1}" -f $PathForError, $lastError)
}

function Acquire-Excel-Automation-Lock {
    param(
        [string]$Name = 'Local\SoftwareContabilidad.ExcelAutomation',
        [int]$TimeoutSeconds = 240
    )

    $mutex = New-Object System.Threading.Mutex($false, $Name)
    $acquired = $false
    try {
        $acquired = $mutex.WaitOne([Math]::Max(1, $TimeoutSeconds) * 1000)
    } catch {
        $mutex.Dispose()
        throw "No se pudo inicializar el bloqueo global de Excel: $($_.Exception.Message)"
    }

    if (-not $acquired) {
        $mutex.Dispose()
        throw "Existe otro proceso de Excel ejecutandose. Espera a que termine e intenta de nuevo."
    }

    return $mutex
}

function Release-Excel-Automation-Lock {
    param([object]$Mutex)

    if ($null -eq $Mutex) {
        return
    }

    try { $Mutex.ReleaseMutex() } catch {}
    try { $Mutex.Dispose() } catch {}
}

function Get-HeadlessExcelProcesses {
    return @(
        Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue |
            Where-Object { $_.MainWindowHandle -eq 0 }
    )
}

function Stop-OrphanExcelProcesses {
    param([int]$TimeoutSeconds = 20)

    $deadline = [datetime]::UtcNow.AddSeconds([Math]::Max(1, $TimeoutSeconds))
    while ($true) {
        $procs = @(Get-HeadlessExcelProcesses)
        if ($procs.Count -eq 0) { return }

        foreach ($proc in $procs) {
            try {
                Stop-Process -Id $proc.Id -Force -ErrorAction Stop
            } catch {
            }
        }

        Start-Sleep -Milliseconds 750

        if ([datetime]::UtcNow -ge $deadline) {
            $pids = ($procs | Select-Object -ExpandProperty Id) -join ','
            Write-Output ("WARN|excel_orphans|Persisten procesos Excel huerfanos (PID: {0}). Se continua con nueva instancia." -f $pids)
            return
        }
    }
}

function Assert-NoVisibleExcelConflict {
    param(
        [string[]]$ConflictPaths = @()
    )

    $visibleProcs = @(
        Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue |
            Where-Object { $_.MainWindowHandle -ne 0 }
    )
    if ($visibleProcs.Count -eq 0) { return }

    $pids = ($visibleProcs | Select-Object -ExpandProperty Id) -join ','

    # Normalize conflict paths to lowercase for comparison
    $normalizedConflicts = @(
        $ConflictPaths |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            ForEach-Object { $_.ToLowerInvariant().TrimEnd('\', '/') }
    )

    if ($normalizedConflicts.Count -eq 0) {
        Write-Output ("WARN|excel_visible_nocheck|Excel visible (PID: {0}) pero sin rutas de conflicto definidas. Se continua." -f $pids)
        return
    }

    # Try to enumerate open workbooks via COM ROT / GetActiveObject
    $conflictFound = $false
    $conflictName = ''
    try {
        $excelApp = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
        if ($null -ne $excelApp) {
            try {
                $workbooks = $excelApp.Workbooks
                for ($i = 1; $i -le $workbooks.Count; $i++) {
                    $wb = $null
                    try {
                        $wb = $workbooks.Item($i)
                        $wbPath = $wb.FullName.ToLowerInvariant()
                        foreach ($conflict in $normalizedConflicts) {
                            if ($wbPath -eq $conflict -or $wbPath.StartsWith($conflict + '\') -or $wbPath.StartsWith($conflict + '/')) {
                                $conflictFound = $true
                                $conflictName = $wb.FullName
                                break
                            }
                        }
                    } catch {}
                    finally {
                        if ($null -ne $wb) {
                            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
                        }
                    }
                    if ($conflictFound) { break }
                }
            } finally {
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelApp)
            }
        }
    } catch {
        # COM check failed (e.g. no active Excel COM object accessible) — safe to continue
        Write-Output ("WARN|excel_visible_comcheck_failed|Excel visible (PID: {0}) pero no se pudo verificar archivos abiertos via COM. Se continua." -f $pids)
        return
    }

    if ($conflictFound) {
        throw "Excel tiene abierto un archivo en uso por el proceso: '$conflictName'. Cierra ese archivo en Excel y vuelve a intentar."
    }

    Write-Output ("WARN|excel_visible_no_conflict|Excel visible (PID: {0}) pero sin archivos en conflicto abiertos. Se continua." -f $pids)
}

function Register-OleMessageFilter {
    if (-not ("OleMessageFilter" -as [type])) {
        Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

[ComImport, Guid("00000016-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IOleMessageFilter
{
    [PreserveSig]
    int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);

    [PreserveSig]
    int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);

    [PreserveSig]
    int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
}

public class OleMessageFilter : IOleMessageFilter
{
    [DllImport("Ole32.dll")]
    private static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, out IOleMessageFilter oldFilter);

    public static void Register()
    {
        IOleMessageFilter newFilter = new OleMessageFilter();
        IOleMessageFilter oldFilter;
        CoRegisterMessageFilter(newFilter, out oldFilter);
    }

    public static void Revoke()
    {
        IOleMessageFilter oldFilter;
        CoRegisterMessageFilter(null, out oldFilter);
    }

    int IOleMessageFilter.HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo)
    {
        return 0;
    }

    int IOleMessageFilter.RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
    {
        if (dwRejectType == 2)
        {
            return 100;
        }

        return -1;
    }

    int IOleMessageFilter.MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType)
    {
        return 2;
    }
}
"@
    }

    try {
        [OleMessageFilter]::Register()
    } catch {
        Write-Output ("WARN|ole_filter|{0}" -f $_.Exception.Message)
    }
}

function Unregister-OleMessageFilter {
    try {
        if ("OleMessageFilter" -as [type]) {
            [OleMessageFilter]::Revoke()
        }
    } catch {
    }
}

function Stop-Excel-Application {
    param(
        [object]$Excel,
        [int]$Attempts = 6
    )

    if ($null -eq $Excel) {
        return
    }

    $lastError = ''
    for ($attempt = 1; $attempt -le $Attempts; $attempt++) {
        try {
            $Excel.Quit()
            return
        } catch {
            $lastError = $_.Exception.Message
            if ($attempt -lt $Attempts) {
                Start-Sleep -Milliseconds (450 * $attempt)
                continue
            }
        }
    }

    if ($lastError -match '0x800AC472|0x80010001') {
        Write-Output ("WARN|excel_quit|{0}" -f $lastError)
        return
    }

    if ($lastError -ne '') {
        Write-Output ("WARN|excel_quit|{0}" -f $lastError)
    }
}

if (-not ('ExcelWindowInterop' -as [type])) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class ExcelWindowInterop {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
"@
}

function Get-ExcelProcessId {
    param([object]$Excel)

    if ($null -eq $Excel) {
        return $null
    }

    try {
        $hwnd = [IntPtr]$Excel.Hwnd
        if ($hwnd -eq [IntPtr]::Zero) {
            return $null
        }

        [uint32]$processId = 0
        [void][ExcelWindowInterop]::GetWindowThreadProcessId($hwnd, [ref]$processId)
        if ($processId -le 0) {
            return $null
        }

        return [int]$processId
    } catch {
        return $null
    }
}

function Stop-WorkerExcelProcess {
    param(
        [Nullable[int]]$ProcessId,
        [int]$WaitMilliseconds = 1200
    )

    if (-not $ProcessId.HasValue -or $ProcessId.Value -le 0) {
        return
    }

    if ($WaitMilliseconds -gt 0) {
        Start-Sleep -Milliseconds $WaitMilliseconds
    }

    $proc = $null
    try {
        $proc = Get-Process -Id $ProcessId.Value -ErrorAction Stop
    } catch {
        return
    }

    try {
        if ($proc.MainWindowHandle -eq 0) {
            Stop-Process -Id $proc.Id -Force -ErrorAction Stop
            Write-Output ("INFO|excel_worker_terminated|pid={0}" -f $proc.Id)
        } else {
            Write-Output ("WARN|excel_worker_visible|pid={0}" -f $proc.Id)
        }
    } catch {
        Write-Output ("WARN|excel_worker_terminate_failed|pid={0}|{1}" -f $ProcessId.Value, $_.Exception.Message)
    }
}

function Normalize-Sheet-Name {
    param([object]$Value)

    $text = (Normalize-Text $Value).ToUpperInvariant()
    $normalized = $text.Normalize([Text.NormalizationForm]::FormD)
    $builder = New-Object System.Text.StringBuilder

    foreach ($char in $normalized.ToCharArray()) {
        $category = [Globalization.CharUnicodeInfo]::GetUnicodeCategory($char)
        if ($category -eq [Globalization.UnicodeCategory]::NonSpacingMark) {
            continue
        }

        if ($char -match '[A-Z0-9 ]') {
            [void]$builder.Append($char)
        }
    }

    return $builder.ToString()
}

function Get-Worksheet-Safe {
    param(
        [object]$Workbook,
        [string[]]$CandidateNames
    )

    $normalizedCandidates = @($CandidateNames | ForEach-Object { Normalize-Sheet-Name $_ })
    foreach ($worksheet in $Workbook.Worksheets) {
        if ($normalizedCandidates -contains (Normalize-Sheet-Name $worksheet.Name)) {
            return $worksheet
        }
    }

    throw "No se encontro la hoja requerida: $($CandidateNames -join ', ')"
}

function To-Number {
    param([object]$Value)

    if ($null -eq $Value -or $Value -eq '') {
        return [double]0
    }

    try {
        return [double]$Value
    } catch {
        $text = Normalize-Text $Value
        if ($text -eq '') {
            return [double]0
        }

        $text = $text.Replace('.', '').Replace(',', '.')
        try {
            return [double]$text
        } catch {
            return [double]0
        }
    }
}

function Round-Amount {
    param([double]$Value)

    return [Math]::Round($Value, 2, [MidpointRounding]::AwayFromZero)
}

function Trim-Document {
    param([object]$Value)

    $text = Normalize-Text $Value
    if ($text -eq '') {
        return ''
    }

    $digits = ($text -replace '[^0-9]', '')
    if ($digits -eq '') {
        return $text
    }

    $trimmed = $digits.TrimStart('0')
    if ($trimmed -eq '') {
        return '0'
    }

    return $trimmed
}

function Strip-Order-Suffix {
    param([object]$Value)

    $text = Normalize-Text $Value
    if ($text -match '^[A-Za-z0-9]+[A-Z]$') {
        return $text.Substring(0, $text.Length - 1)
    }

    return $text
}

function Get-Date-Value {
    param(
        [object]$Worksheet,
        [int]$Row,
        [int]$Column
    )

    $raw = $Worksheet.Cells.Item($Row, $Column).Value2
    if ($raw -is [double] -or $raw -is [int] -or $raw -is [decimal]) {
        return [double]$raw
    }

    $text = Normalize-Text $Worksheet.Cells.Item($Row, $Column).Text
    if ($text -eq '') {
        return $null
    }

    try {
        return [double]([datetime]::Parse($text).ToOADate())
    } catch {
        return $null
    }
}

function Get-Date-Write-Value {
    param([object]$Value)

    if ($null -eq $Value -or $Value -eq '') {
        return $null
    }

    return [int][Math]::Floor(([double]$Value) + 0.000001)
}

function Set-CellValue {
    param(
        [object]$Worksheet,
        [int]$Row,
        [int]$Column,
        [object]$Value,
        [string]$Context = ''
    )

    try {
        $Worksheet.Cells.Item($Row, $Column).Value2 = $Value
    } catch {
        $valueType = if ($null -eq $Value) { 'null' } else { $Value.GetType().FullName }
        throw "No se pudo escribir [$Context] row=$Row col=$Column type=$valueType value=$Value :: $($_.Exception.Message)"
    }
}

function Set-DateCellValue {
    param(
        [object]$Worksheet,
        [int]$Row,
        [int]$Column,
        [object]$Value,
        [string]$Context = ''
    )

    $cell = $Worksheet.Cells.Item($Row, $Column)
    try {
        if ($null -eq $Value -or $Value -eq '') {
            $null = $cell.ClearContents()
            return
        }

        $dateSerial = [double]$Value
        $currentFormat = Normalize-Text $cell.NumberFormat
        $dateText = [datetime]::FromOADate($dateSerial).ToString('dd/MM/yyyy')

        if ($currentFormat -eq '' -or $currentFormat -eq 'General') {
            # En formato General Excel cambia el NumberFormat al asignar DateTime.
            # Escribimos texto para conservar exactamente el formato de la plantilla manual.
            $null = $cell.NumberFormat = 'General'
            $null = $cell.Value2 = "'" + $dateText
        } else {
            # Para celdas ya formateadas como fecha, mantener formato y grabar serial.
            $null = $cell.Value = [datetime]::FromOADate($dateSerial)
        }
    } catch {
        $valueType = if ($null -eq $Value) { 'null' } else { $Value.GetType().FullName }
        throw "No se pudo escribir fecha [$Context] row=$Row col=$Column type=$valueType value=$Value :: $($_.Exception.Message)"
    } finally {
        [void][Runtime.Interopservices.Marshal]::ReleaseComObject($cell)
    }
}

function Validate-Services-SourceWorksheet {
    param([object]$Worksheet)

    $checks = @(
        @{ Row = 1; Column = 1; Needle = 'AGENCIA' },
        @{ Row = 1; Column = 2; Needle = 'CENTRO' },
        @{ Row = 1; Column = 3; Needle = 'No. ORDEN' },
        @{ Row = 1; Column = 8; Needle = 'TIPO DOC' },
        @{ Row = 1; Column = 9; Needle = 'CEDULA' },
        @{ Row = 1; Column = 10; Needle = 'FACTURADO A' },
        @{ Row = 1; Column = 12; Needle = 'DOCUMENTO' },
        @{ Row = 1; Column = 15; Needle = 'F. FACT' },
        @{ Row = 1; Column = 18; Needle = 'F. NOTA' },
        @{ Row = 1; Column = 36; Needle = 'ANULADA' }
    )

    foreach ($check in $checks) {
        $actual = (Normalize-Text $Worksheet.Cells.Item($check.Row, $check.Column).Text).ToUpperInvariant()
        $expected = (Normalize-Text $check.Needle).ToUpperInvariant()
        if ($actual -notlike "*$expected*") {
            throw ("El archivo fuente no coincide con la estructura esperada en fila {0} columna {1}. Esperado contiene '{2}' y llego '{3}'." -f $check.Row, $check.Column, $check.Needle, $actual)
        }
    }
}

function Get-ComparableCellText {
    param(
        [object]$Worksheet,
        [int]$Row,
        [int]$Column
    )

    $value = $Worksheet.Cells.Item($Row, $Column).Value2
    if ($null -ne $value -and $value -ne '') {
        return Normalize-ExcelLogicalText $value
    }

    return Normalize-ExcelLogicalText $Worksheet.Cells.Item($Row, $Column).Text
}

function Normalize-ExcelLogicalText {
    param([object]$Value)

    $text = Normalize-Text $Value
    if ($text -ne '') {
        $text = ($text -replace '\s+', ' ').Trim()
    }

    if ($text.Length -gt 1 -and $text.StartsWith("'")) {
        $text = $text.Substring(1)
    }

    if ($text -ne '') {
        $text = ($text -replace '\s+', ' ').Trim()
    }

    return $text
}

function Get-ComparableCellNumber {
    param(
        [object]$Worksheet,
        [int]$Row,
        [int]$Column
    )

    $raw = $Worksheet.Cells.Item($Row, $Column).Value2
    if ($null -eq $raw -or $raw -eq '') {
        $text = Normalize-Text $Worksheet.Cells.Item($Row, $Column).Text
        if ($text -eq '') {
            return $null
        }

        $normalized = ($text -replace '\.', '') -replace ',', '.'
        return [double]::Parse($normalized, [System.Globalization.CultureInfo]::InvariantCulture)
    }

    return [double]$raw
}

function Get-WorksheetRowPropertyValues {
    param(
        [object]$Worksheet,
        [int]$Row,
        [int]$LastColumn,
        [string]$PropertyName
    )

    if ($null -eq $Worksheet -or $Row -lt 1 -or $LastColumn -lt 1) {
        return $null
    }

    $startCell = $null
    $endCell = $null
    $range = $null
    try {
        $startCell = $Worksheet.Cells.Item($Row, 1)
        $endCell = $Worksheet.Cells.Item($Row, $LastColumn)
        $range = $Worksheet.Range($startCell, $endCell)
        return $range.$PropertyName
    } finally {
        if ($null -ne $range) {
            [void][Runtime.Interopservices.Marshal]::ReleaseComObject($range)
        }
        if ($null -ne $endCell) {
            [void][Runtime.Interopservices.Marshal]::ReleaseComObject($endCell)
        }
        if ($null -ne $startCell) {
            [void][Runtime.Interopservices.Marshal]::ReleaseComObject($startCell)
        }
    }
}

function Get-WorksheetRangePropertyValues {
    param(
        [object]$Worksheet,
        [int]$StartRow,
        [int]$EndRow,
        [int]$LastColumn,
        [string]$PropertyName
    )

    if ($null -eq $Worksheet -or $StartRow -lt 1 -or $EndRow -lt $StartRow -or $LastColumn -lt 1) {
        return $null
    }

    $startCell = $null
    $endCell = $null
    $range = $null
    try {
        $startCell = $Worksheet.Cells.Item($StartRow, 1)
        $endCell = $Worksheet.Cells.Item($EndRow, $LastColumn)
        $range = $Worksheet.Range($startCell, $endCell)
        return $range.$PropertyName
    } finally {
        if ($null -ne $range) {
            [void][Runtime.Interopservices.Marshal]::ReleaseComObject($range)
        }
        if ($null -ne $endCell) {
            [void][Runtime.Interopservices.Marshal]::ReleaseComObject($endCell)
        }
        if ($null -ne $startCell) {
            [void][Runtime.Interopservices.Marshal]::ReleaseComObject($startCell)
        }
    }
}

function Get-WorksheetRowMatrixValue {
    param(
        [object]$Matrix,
        [int]$Column
    )

    if ($null -eq $Matrix) {
        return $null
    }

    if ($Matrix -is [System.Array]) {
        if ($Matrix.Rank -ge 2) {
            return $Matrix.GetValue(1, $Column)
        }

        return $Matrix.GetValue($Column - 1)
    }

    return $Matrix
}

function Get-WorksheetRangeMatrixValue {
    param(
        [object]$Matrix,
        [int]$RowOffset,
        [int]$Column
    )

    if ($null -eq $Matrix) {
        return $null
    }

    if ($Matrix -is [System.Array]) {
        if ($Matrix.Rank -ge 2) {
            return $Matrix.GetValue($RowOffset, $Column)
        }

        if ($RowOffset -eq 1) {
            return $Matrix.GetValue($Column - 1)
        }

        return $null
    }

    if ($RowOffset -eq 1 -and $Column -eq 1) {
        return $Matrix
    }

    return $null
}

function Test-NumericEquivalent {
    param(
        [Nullable[double]]$Expected,
        [Nullable[double]]$Actual
    )

    if (-not $Expected.HasValue -and -not $Actual.HasValue) {
        return $true
    }

    if (-not $Expected.HasValue -or -not $Actual.HasValue) {
        return $false
    }

    return ([Math]::Abs([double]$Expected.Value - [double]$Actual.Value) -lt 0.01)
}

function Assert-NoExcelErrorsInRange {
    param(
        [object]$Worksheet,
        [int]$StartRow,
        [int]$EndRow,
        [int]$LastColumn,
        [string]$Label
    )

    if ($null -eq $Worksheet -or $StartRow -gt $EndRow -or $LastColumn -lt 1) {
        return
    }

    $startCell = $null
    $endCell = $null
    $range = $null
    $errorRange = $null
    $firstCell = $null
    try {
        $startCell = $Worksheet.Cells.Item($StartRow, 1)
        $endCell = $Worksheet.Cells.Item($EndRow, $LastColumn)
        $range = $Worksheet.Range($startCell, $endCell)

        foreach ($cellType in @(-4123, 2)) {
            Assert-NotCancelled 'validacion_errores'
            try {
                $errorRange = $range.SpecialCells($cellType, 16)
            } catch {
                $errorRange = $null
            }

            if ($null -eq $errorRange) {
                continue
            }

            try {
                $firstCell = $errorRange.Areas.Item(1).Cells.Item(1, 1)
                $text = (Normalize-Text $firstCell.Text).ToUpperInvariant()
                throw ("La hoja {0} contiene error de Excel en fila {1} columna {2}: {3}" -f $Label, [int]$firstCell.Row, [int]$firstCell.Column, $text)
            } finally {
                if ($null -ne $firstCell) {
                    [void][Runtime.Interopservices.Marshal]::ReleaseComObject($firstCell)
                    $firstCell = $null
                }
                if ($null -ne $errorRange) {
                    [void][Runtime.Interopservices.Marshal]::ReleaseComObject($errorRange)
                    $errorRange = $null
                }
            }
        }
    } finally {
        if ($null -ne $range) {
            [void][Runtime.Interopservices.Marshal]::ReleaseComObject($range)
        }
        if ($null -ne $endCell) {
            [void][Runtime.Interopservices.Marshal]::ReleaseComObject($endCell)
        }
        if ($null -ne $startCell) {
            [void][Runtime.Interopservices.Marshal]::ReleaseComObject($startCell)
        }
    }
}

function Assert-WorksheetAreaBlank {
    param(
        [object]$Worksheet,
        [int]$StartRow,
        [int]$EndRow,
        [int]$StartColumn,
        [int]$EndColumn,
        [string]$Label
    )

    if ($null -eq $Worksheet -or $StartRow -gt $EndRow -or $StartColumn -gt $EndColumn) {
        return
    }

    $startCell = $null
    $endCell = $null
    $range = $null
    try {
        $startCell = $Worksheet.Cells.Item($StartRow, $StartColumn)
        $endCell = $Worksheet.Cells.Item($EndRow, $EndColumn)
        $range = $Worksheet.Range($startCell, $endCell)
        $nonBlankCount = [int]$Worksheet.Application.WorksheetFunction.CountA($range)
        if ($nonBlankCount -le 0) {
            return
        }
    } finally {
        if ($null -ne $range) {
            [void][Runtime.Interopservices.Marshal]::ReleaseComObject($range)
        }
        if ($null -ne $endCell) {
            [void][Runtime.Interopservices.Marshal]::ReleaseComObject($endCell)
        }
        if ($null -ne $startCell) {
            [void][Runtime.Interopservices.Marshal]::ReleaseComObject($startCell)
        }
    }

    for ($row = $StartRow; $row -le $EndRow; $row++) {
        Assert-NotCancelled 'validacion_rango_en_blanco'
        for ($column = $StartColumn; $column -le $EndColumn; $column++) {
            $text = Normalize-Text $Worksheet.Cells.Item($row, $column).Text
            if ($text -ne '') {
                throw ("La hoja {0} conserva dato historico o residual en fila {1} columna {2}: '{3}'." -f $Label, $row, $column, $text)
            }
        }
    }
}

function Validate-Services-WrittenRows {
    param(
        [object]$OutputWorksheet,
        [object]$TemplateWorksheet,
        [object[]]$Rows,
        [string]$Label,
        [int[]]$TextColumns,
        [int[]]$NumericColumns,
        [int[]]$DateColumns,
        [int[]]$FormatColumns,
        [int[]]$DocumentColumns = @(),
        [int]$TrailingBlankColumn,
        [int]$LastColumn
    )

    if ($Rows.Count -eq 0) {
        $firstCell = Normalize-Text $OutputWorksheet.Cells.Item(1, $TrailingBlankColumn).Text
        if ($firstCell -in @('#DIV/0!', '#N/A', '#NAME?', '#NULL!', '#NUM!', '#REF!', '#VALUE!')) {
            throw "La hoja $Label contiene errores aun sin filas escritas."
        }
        return
    }

    $startRow = [int]$Rows[0].RowNumber
    $endRow = [int]$Rows[-1].RowNumber
    $outputValues = Get-WorksheetRangePropertyValues -Worksheet $OutputWorksheet -StartRow $startRow -EndRow $endRow -LastColumn $LastColumn -PropertyName 'Value2'
    $outputFormats = $null
    $templateFormats = $null
    if (@($FormatColumns).Count -gt 0) {
        $outputFormats = Get-WorksheetRangePropertyValues -Worksheet $OutputWorksheet -StartRow $startRow -EndRow $endRow -LastColumn $LastColumn -PropertyName 'NumberFormat'
        $templateFormats = Get-WorksheetRangePropertyValues -Worksheet $TemplateWorksheet -StartRow $startRow -EndRow $endRow -LastColumn $LastColumn -PropertyName 'NumberFormat'
    }

    foreach ($entry in $Rows) {
        Assert-NotCancelled 'validacion_filas'
        $rowNumber = [int]$entry.RowNumber
        $rowOffset = [int]($rowNumber - $startRow + 1)
        $values = $entry.Values

        foreach ($column in $TextColumns) {
            $expected = if ($values.ContainsKey($column)) { Normalize-ExcelLogicalText $values[$column] } else { '' }
            $actual = Normalize-ExcelLogicalText (Get-WorksheetRangeMatrixValue -Matrix $outputValues -RowOffset $rowOffset -Column $column)
            if ($column -in $DocumentColumns) {
                $expected = Trim-Document $expected
                $actual = Trim-Document $actual
            }
            if ($actual -ne $expected) {
                throw ("La hoja {0} tiene texto incorrecto en fila {1} columna {2}. Esperado '{3}' y llego '{4}'." -f $Label, $rowNumber, $column, $expected, $actual)
            }
        }

        foreach ($column in $NumericColumns) {
            $expectedValue = $null
            if ($values.ContainsKey($column) -and $null -ne $values[$column] -and $values[$column] -ne '') {
                $expectedValue = [double]$values[$column]
            }

            $actualRaw = Get-WorksheetRangeMatrixValue -Matrix $outputValues -RowOffset $rowOffset -Column $column
            $actualValue = $null
            $actualText = Normalize-Text $actualRaw
            if ($null -ne $actualRaw -and $actualRaw -ne '') {
                if ($actualRaw -is [double] -or $actualRaw -is [int] -or $actualRaw -is [decimal]) {
                    $actualValue = [double]$actualRaw
                } else {
                    $normalized = ($actualText -replace '\.', '') -replace ',', '.'
                    if ($normalized -ne '') {
                        try {
                            $actualValue = [double]::Parse($normalized, [System.Globalization.CultureInfo]::InvariantCulture)
                        } catch {
                            $actualValue = $null
                        }
                    }
                }
            }

            if ($null -eq $expectedValue -and $actualText -eq '') {
                continue
            }

            if (-not (Test-NumericEquivalent -Expected $expectedValue -Actual $actualValue)) {
                throw ("La hoja {0} tiene valor numerico incorrecto en fila {1} columna {2}. Esperado '{3}' y llego '{4}'." -f $Label, $rowNumber, $column, $expectedValue, $actualValue)
            }
        }

        foreach ($column in $DateColumns) {
            $expectedDate = $null
            if ($values.ContainsKey($column) -and $null -ne $values[$column] -and $values[$column] -ne '') {
                $expectedDate = [double](Get-Date-Write-Value $values[$column])
            }

            $actualDate = $null
            $actualRaw = Get-WorksheetRangeMatrixValue -Matrix $outputValues -RowOffset $rowOffset -Column $column
            if ($null -ne $actualRaw -and $actualRaw -ne '') {
                if ($actualRaw -is [double] -or $actualRaw -is [int] -or $actualRaw -is [decimal]) {
                    $actualDate = [double](Get-Date-Write-Value $actualRaw)
                } else {
                    try {
                        $actualDate = [double]([datetime]::Parse((Normalize-Text $actualRaw)).ToOADate())
                        $actualDate = [double](Get-Date-Write-Value $actualDate)
                    } catch {
                        $actualDate = $null
                    }
                }
            }
            if ($null -eq $expectedDate -and $null -eq $actualDate) {
                continue
            }

            if ($null -eq $expectedDate -or $null -eq $actualDate -or [Math]::Abs([double]$expectedDate - [double]$actualDate) -ge 0.01) {
                throw ("La hoja {0} tiene fecha incorrecta en fila {1} columna {2}. Esperado '{3}' y llego '{4}'." -f $Label, $rowNumber, $column, $expectedDate, $actualDate)
            }
        }

        foreach ($column in $FormatColumns) {
            $templateFormat = Normalize-Text (Get-WorksheetRangeMatrixValue -Matrix $templateFormats -RowOffset $rowOffset -Column $column)
            $outputFormat = Normalize-Text (Get-WorksheetRangeMatrixValue -Matrix $outputFormats -RowOffset $rowOffset -Column $column)
            if ($templateFormat -ne $outputFormat) {
                throw ("La hoja {0} perdio formato en fila {1} columna {2}. Esperado '{3}' y llego '{4}'." -f $Label, $rowNumber, $column, $templateFormat, $outputFormat)
            }
        }
    }

    $nextRow = $endRow + 1
    if ((Normalize-Text $OutputWorksheet.Cells.Item($nextRow, $TrailingBlankColumn).Text) -ne '') {
        throw ("La hoja {0} conserva datos residuales despues de la ultima fila esperada ({1})." -f $Label, $Rows[-1].RowNumber)
    }

    Assert-NoExcelErrorsInRange -Worksheet $OutputWorksheet -StartRow $startRow -EndRow $endRow -LastColumn $LastColumn -Label $Label
}

function Get-Template-Key {
    param(
        [object]$Agency,
        [object]$Series,
        [object]$Order = ''
    )

    $agencyText = (Normalize-Text $Agency).ToUpperInvariant()
    $seriesText = Normalize-Text $Series
    $orderText = (Normalize-Text $Order).ToUpperInvariant()

    switch ($agencyText) {
        'CHANGAN' { return 'changan' }
        'PEUGEOT' { return 'peug' }
        'MATRIZ' { return 'tyt' }
        'SUZUKI AMBATO' { return 'szk' }
        'SUZUKI RIOBAMBA' {
            # Regla operativa del formato manual: ciertos documentos D.... van en CHANGAN.
            if ($orderText -match '^D\d+') {
                return 'changan'
            }

            return 'szk'
        }
        default { return $null }
    }
}

function Get-Document-SortValue {
    param([object]$Value)

    $text = Trim-Document $Value
    if ($text -match '^\d+$') {
        return [double]$text
    }

    return [double]0
}

function New-LookupStore {
    return @{
        ByDocOrder = @{}
        ByDoc = @{}
        ByOrder = @{}
    }
}

function Get-FileSignature {
    param([string]$Path)

    $item = Get-Item -LiteralPath $Path
    return [pscustomobject]@{
        Path = $item.FullName
        Length = [int64]$item.Length
        LastWriteTimeUtc = $item.LastWriteTimeUtc.ToString('o')
    }
}

function Get-TemplateSnapshotCachePath {
    param([string]$TemplatePath)

    $cacheRoot = Join-Path $PSScriptRoot 'storage\cache\servicios_marcas'
    if (-not (Test-Path -LiteralPath $cacheRoot)) {
        New-Item -ItemType Directory -Path $cacheRoot -Force | Out-Null
    }

    $resolvedPath = (Resolve-Path -LiteralPath $TemplatePath).Path
    $sha1 = [System.Security.Cryptography.SHA1]::Create()
    try {
        $hashBytes = $sha1.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($resolvedPath))
    } finally {
        $sha1.Dispose()
    }

    $hash = [System.BitConverter]::ToString($hashBytes).Replace('-', '').ToLowerInvariant()
    return (Join-Path $cacheRoot ("template_snapshot_{0}.clixml" -f $hash))
}

function Test-TemplateSnapshotSignature {
    param(
        [object]$SnapshotSignature,
        [object]$CurrentSignature
    )

    if ($null -eq $SnapshotSignature -or $null -eq $CurrentSignature) {
        return $false
    }

    return (
        ([string]$SnapshotSignature.Path -eq [string]$CurrentSignature.Path) -and
        ([string]$SnapshotSignature.Length -eq [string]$CurrentSignature.Length) -and
        ([string]$SnapshotSignature.LastWriteTimeUtc -eq [string]$CurrentSignature.LastWriteTimeUtc)
    )
}

function Get-TemplateSnapshot {
    param(
        [object]$Workbook,
        [string]$TemplatePath,
        [string]$TemplateKey
    )

    $signature = Get-FileSignature -Path $TemplatePath
    $cachePath = Get-TemplateSnapshotCachePath -TemplatePath $TemplatePath
    if (Test-Path -LiteralPath $cachePath) {
        try {
            $cachedRecord = Import-Clixml -LiteralPath $cachePath
            if (Test-TemplateSnapshotSignature -SnapshotSignature $cachedRecord.Signature -CurrentSignature $signature) {
                Write-Output ("INFO|template_snapshot_cache|{0}|hit" -f $TemplateKey)
                return $cachedRecord.Data
            }

            Write-Output ("INFO|template_snapshot_cache|{0}|stale" -f $TemplateKey)
        } catch {
            Write-Output ("WARN|template_snapshot_cache_read_failed|{0}|{1}" -f $TemplateKey, $_.Exception.Message)
        }
    } else {
        Write-Output ("INFO|template_snapshot_cache|{0}|miss" -f $TemplateKey)
    }

    $lookupStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $lookups = Read-TemplateLookups -Workbook $Workbook
    $lookupStopwatch.Stop()
    Write-Output ("INFO|template_lookup_ms|{0}|{1}" -f $TemplateKey, $lookupStopwatch.ElapsedMilliseconds)

    $templatePrecontVentasSheet = $null
    $templatePrecontCostos2Sheet = $null
    $precontVentasPrototypes = @()
    $precontCostos2Prototypes = @()
    try {
        try { $templatePrecontVentasSheet = Get-Worksheet-Safe -Workbook $Workbook -CandidateNames @('PrecontabilizacionVentas') } catch {}
        if ($null -ne $templatePrecontVentasSheet) {
            $ventasProtoStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            $precontVentasPrototypes = @(Read-PrecontVentasPrototypes -Worksheet $templatePrecontVentasSheet)
            $ventasProtoStopwatch.Stop()
            Write-Output ("INFO|template_precont_ventas_proto_ms|{0}|{1}" -f $TemplateKey, $ventasProtoStopwatch.ElapsedMilliseconds)
        }

        try { $templatePrecontCostos2Sheet = Get-Worksheet-Safe -Workbook $Workbook -CandidateNames @('PrecontabilizacionCostos (2)') } catch {}
        if ($null -ne $templatePrecontCostos2Sheet) {
            $costosProtoStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            $precontCostos2Prototypes = @(Read-PrecontCostos2Prototypes -Worksheet $templatePrecontCostos2Sheet)
            $costosProtoStopwatch.Stop()
            Write-Output ("INFO|template_precont_costos2_proto_ms|{0}|{1}" -f $TemplateKey, $costosProtoStopwatch.ElapsedMilliseconds)
        }
    } finally {
        if ($null -ne $templatePrecontCostos2Sheet) {
            [void][Runtime.Interopservices.Marshal]::ReleaseComObject($templatePrecontCostos2Sheet)
        }
        if ($null -ne $templatePrecontVentasSheet) {
            [void][Runtime.Interopservices.Marshal]::ReleaseComObject($templatePrecontVentasSheet)
        }
    }

    $snapshot = [pscustomobject]@{
        Lookups = $lookups
        PrecontVentasPrototypes = @($precontVentasPrototypes)
        PrecontCostos2Prototypes = @($precontCostos2Prototypes)
    }

    try {
        [pscustomobject]@{
            Signature = $signature
            Data = $snapshot
        } | Export-Clixml -LiteralPath $cachePath -Depth 8 -Force
        Write-Output ("INFO|template_snapshot_cache|{0}|write" -f $TemplateKey)
    } catch {
        Write-Output ("WARN|template_snapshot_cache_write_failed|{0}|{1}" -f $TemplateKey, $_.Exception.Message)
    }

    return $snapshot
}

function Add-LookupEntry {
    param(
        [hashtable]$Store,
        [object]$DocKey,
        [object]$OrderKey,
        [pscustomobject]$Entry
    )

    if ($DocKey -ne '' -and $OrderKey -ne '') {
        $docOrderKey = $DocKey + '|' + $OrderKey
        if (-not $Store.ByDocOrder.ContainsKey($docOrderKey)) {
            $Store.ByDocOrder[$docOrderKey] = $Entry
        }
    }

    if ($DocKey -ne '' -and -not $Store.ByDoc.ContainsKey($DocKey)) {
        $Store.ByDoc[$DocKey] = $Entry
    }

    if ($OrderKey -ne '' -and -not $Store.ByOrder.ContainsKey($OrderKey)) {
        $Store.ByOrder[$OrderKey] = $Entry
    }
}

function Find-LookupEntry {
    param(
        [hashtable]$Store,
        [object]$DocKey,
        [object]$OrderKey
    )

    if ($DocKey -ne '' -and $OrderKey -ne '') {
        $docOrderKey = $DocKey + '|' + $OrderKey
        if ($Store.ByDocOrder.ContainsKey($docOrderKey)) {
            return $Store.ByDocOrder[$docOrderKey]
        }
    }

    if ($DocKey -ne '' -and $Store.ByDoc.ContainsKey($DocKey)) {
        return $Store.ByDoc[$DocKey]
    }

    if ($OrderKey -ne '' -and $Store.ByOrder.ContainsKey($OrderKey)) {
        return $Store.ByOrder[$OrderKey]
    }

    return $null
}

function Find-RepVtasEntry {
    param(
        [hashtable]$Store,
        [object]$DocKey,
        [object]$OrderKey
    )

    if ($DocKey -ne '' -and $OrderKey -ne '') {
        $docOrderKey = $DocKey + '|' + $OrderKey
        if ($Store.ByDocOrder.ContainsKey($docOrderKey)) {
            return $Store.ByDocOrder[$docOrderKey]
        }
    }

    if ($OrderKey -ne '' -and $Store.ByOrder.ContainsKey($OrderKey)) {
        return $Store.ByOrder[$OrderKey]
    }

    if ($DocKey -ne '' -and $Store.ByDoc.ContainsKey($DocKey)) {
        return $Store.ByDoc[$DocKey]
    }

    return $null
}

function Get-Worksheet-Column-DefaultText {
    param(
        [object]$Worksheet,
        [int]$StartRow,
        [int]$KeyColumn,
        [int]$ValueColumn
    )

    $lastRow = $Worksheet.Cells.Item($Worksheet.Rows.Count, $KeyColumn).End(-4162).Row
    for ($row = $StartRow; $row -le $lastRow; $row++) {
        $keyText = Normalize-Text $Worksheet.Cells.Item($row, $KeyColumn).Text
        if ($keyText -eq '') {
            continue
        }

        $valueText = Normalize-Text $Worksheet.Cells.Item($row, $ValueColumn).Text
        if ($valueText -ne '') {
            return $valueText
        }
    }

    return ''
}

function Read-TemplateLookups {
    param([object]$Workbook)

    Assert-NotCancelled 'lookup'
    $invoiceStore = New-LookupStore
    $noteStore = New-LookupStore
    $repVtasStore = New-LookupStore

    $invoiceSheet = Get-Worksheet-Safe -Workbook $Workbook -CandidateNames @('REP FACTURACION', 'REP FACTURACIÓN')
    $noteSheet = Get-Worksheet-Safe -Workbook $Workbook -CandidateNames @('NOTA DE CREDITO')

    $repVtasSheet = Get-Worksheet-Safe -Workbook $Workbook -CandidateNames @('REP VTAS')
    $defaults = @{
        Invoice = @{
            GarExt = Get-Worksheet-Column-DefaultText -Worksheet $invoiceSheet -StartRow 17 -KeyColumn 3 -ValueColumn 17
        }
        Note = @{
            GarExt = Get-Worksheet-Column-DefaultText -Worksheet $noteSheet -StartRow 11 -KeyColumn 2 -ValueColumn 21
        }
        RepVtas = @{
            GarExt = Get-Worksheet-Column-DefaultText -Worksheet $repVtasSheet -StartRow 15 -KeyColumn 8 -ValueColumn 27
        }
    }
    $lastInvoiceRow = $invoiceSheet.Cells.Item($invoiceSheet.Rows.Count, 3).End(-4162).Row
    for ($row = 17; $row -le $lastInvoiceRow; $row++) {
        Assert-NotCancelled 'lookup_facturas'
        $docKey = Trim-Document $invoiceSheet.Cells.Item($row, 3).Text
        if ($docKey -eq '') {
            continue
        }

        $orderKey = Normalize-Text $invoiceSheet.Cells.Item($row, 5).Text
        $entry = [pscustomobject]@{
            Agency = Normalize-Text $invoiceSheet.Cells.Item($row, 1).Text
            Series = Normalize-Text $invoiceSheet.Cells.Item($row, 2).Text
            Document = $docKey
            Order = $orderKey
            Cedula = Normalize-Text $invoiceSheet.Cells.Item($row, 6).Text
            Customer = Normalize-Text $invoiceSheet.Cells.Item($row, 7).Text
            Subtotal = To-Number $invoiceSheet.Cells.Item($row, 8).Value2
            Discount = To-Number $invoiceSheet.Cells.Item($row, 9).Value2
            NetoConIva = To-Number $invoiceSheet.Cells.Item($row, 10).Value2
            NetoIva0 = To-Number $invoiceSheet.Cells.Item($row, 11).Value2
            Iva12 = To-Number $invoiceSheet.Cells.Item($row, 12).Value2
            Iva = To-Number $invoiceSheet.Cells.Item($row, 13).Value2
            Interest = To-Number $invoiceSheet.Cells.Item($row, 14).Value2
            Total = To-Number $invoiceSheet.Cells.Item($row, 15).Value2
            IvaText = Normalize-Text $invoiceSheet.Cells.Item($row, 13).Text
            Asiento = Normalize-Text $invoiceSheet.Cells.Item($row, 16).Text
            GarExt = Normalize-Text $invoiceSheet.Cells.Item($row, 17).Text
            Tv = Normalize-Text $invoiceSheet.Cells.Item($row, 18).Text
            Marker = Normalize-Text $invoiceSheet.Cells.Item($row, 19).Text
        }

        Add-LookupEntry -Store $invoiceStore -DocKey $docKey -OrderKey $orderKey -Entry $entry
    }

    $lastNoteRow = $noteSheet.Cells.Item($noteSheet.Rows.Count, 2).End(-4162).Row
    for ($row = 11; $row -le $lastNoteRow; $row++) {
        Assert-NotCancelled 'lookup_notas'
        $docKey = Trim-Document $noteSheet.Cells.Item($row, 2).Text
        if ($docKey -eq '') {
            continue
        }

        $orderKey = Normalize-Text $noteSheet.Cells.Item($row, 7).Text
        $entry = [pscustomobject]@{
            Agency = Normalize-Text $noteSheet.Cells.Item($row, 1).Text
            Document = $docKey
            Kind = Normalize-Text $noteSheet.Cells.Item($row, 4).Text
            Series = Normalize-Text $noteSheet.Cells.Item($row, 5).Text
            Invoice = Trim-Document $noteSheet.Cells.Item($row, 6).Text
            Order = $orderKey
            Cedula = Normalize-Text $noteSheet.Cells.Item($row, 8).Text
            Customer = Normalize-Text $noteSheet.Cells.Item($row, 9).Text
            Subtotal = To-Number $noteSheet.Cells.Item($row, 10).Value2
            Discount = To-Number $noteSheet.Cells.Item($row, 11).Value2
            NetoSinIva = To-Number $noteSheet.Cells.Item($row, 12).Value2
            NetoConIva = To-Number $noteSheet.Cells.Item($row, 13).Value2
            Iva = To-Number $noteSheet.Cells.Item($row, 14).Value2
            Iva12 = To-Number $noteSheet.Cells.Item($row, 15).Value2
            Interest = To-Number $noteSheet.Cells.Item($row, 16).Value2
            Total = To-Number $noteSheet.Cells.Item($row, 17).Value2
            Anticipo = To-Number $noteSheet.Cells.Item($row, 18).Value2
            Neto = To-Number $noteSheet.Cells.Item($row, 19).Value2
            Asiento = Normalize-Text $noteSheet.Cells.Item($row, 20).Text
            GarExt = Normalize-Text $noteSheet.Cells.Item($row, 21).Text
        }

        Add-LookupEntry -Store $noteStore -DocKey $docKey -OrderKey $orderKey -Entry $entry
    }

    $lastRepVtasRow = $repVtasSheet.Cells.Item($repVtasSheet.Rows.Count, 3).End(-4162).Row
    for ($row = 15; $row -le $lastRepVtasRow; $row++) {
        Assert-NotCancelled 'lookup_rep_vtas'
        $docKey = Trim-Document $repVtasSheet.Cells.Item($row, 8).Text
        $orderKey = Normalize-Text $repVtasSheet.Cells.Item($row, 3).Text
        if ($docKey -eq '' -and $orderKey -eq '') {
            continue
        }

        $entry = [pscustomobject]@{
            RowOrder = $row
            Agency = Normalize-Text $repVtasSheet.Cells.Item($row, 1).Text
            Center = Normalize-Text $repVtasSheet.Cells.Item($row, 2).Text
            Order = $orderKey
            Advisor = Normalize-Text $repVtasSheet.Cells.Item($row, 4).Text
            Line = Normalize-Text $repVtasSheet.Cells.Item($row, 5).Text
            Cedula = Normalize-Text $repVtasSheet.Cells.Item($row, 6).Text
            Customer = Normalize-Text $repVtasSheet.Cells.Item($row, 7).Text
            DocumentRaw = Normalize-Text $repVtasSheet.Cells.Item($row, 8).Text
            DateFactValue = Get-Date-Value -Worksheet $repVtasSheet -Row $row -Column 9
            DateNoteValue = Get-Date-Value -Worksheet $repVtasSheet -Row $row -Column 10
            NoteCredit = To-Number $repVtasSheet.Cells.Item($row, 11).Value2
            NoteCreditText = Normalize-Text $repVtasSheet.Cells.Item($row, 11).Text
            TotalManoObra = To-Number $repVtasSheet.Cells.Item($row, 12).Value2
            TotalManoObraText = Normalize-Text $repVtasSheet.Cells.Item($row, 12).Text
            TotalSubcontratos = To-Number $repVtasSheet.Cells.Item($row, 13).Value2
            TotalSubcontratosText = Normalize-Text $repVtasSheet.Cells.Item($row, 13).Text
            TotalInsumos = To-Number $repVtasSheet.Cells.Item($row, 14).Value2
            TotalInsumosText = Normalize-Text $repVtasSheet.Cells.Item($row, 14).Text
            TotalServicio = To-Number $repVtasSheet.Cells.Item($row, 15).Value2
            TotalServicioText = Normalize-Text $repVtasSheet.Cells.Item($row, 15).Text
            TotalAccesorios = To-Number $repVtasSheet.Cells.Item($row, 16).Value2
            TotalAccesoriosText = Normalize-Text $repVtasSheet.Cells.Item($row, 16).Text
            TotalRepuestos = To-Number $repVtasSheet.Cells.Item($row, 17).Value2
            TotalRepuestosText = Normalize-Text $repVtasSheet.Cells.Item($row, 17).Text
            Interes = To-Number $repVtasSheet.Cells.Item($row, 18).Value2
            InteresText = Normalize-Text $repVtasSheet.Cells.Item($row, 18).Text
            Iva = To-Number $repVtasSheet.Cells.Item($row, 19).Value2
            IvaText = Normalize-Text $repVtasSheet.Cells.Item($row, 19).Text
            Total = To-Number $repVtasSheet.Cells.Item($row, 20).Value2
            TotalText = Normalize-Text $repVtasSheet.Cells.Item($row, 20).Text
            Costo = To-Number $repVtasSheet.Cells.Item($row, 21).Value2
            CostoText = Normalize-Text $repVtasSheet.Cells.Item($row, 21).Text
            CostoLubricantes = To-Number $repVtasSheet.Cells.Item($row, 22).Value2
            CostoLubricantesText = Normalize-Text $repVtasSheet.Cells.Item($row, 22).Text
            CostoAccesorios = To-Number $repVtasSheet.Cells.Item($row, 23).Value2
            CostoAccesoriosText = Normalize-Text $repVtasSheet.Cells.Item($row, 23).Text
            CostoRepuestos = To-Number $repVtasSheet.Cells.Item($row, 24).Value2
            CostoRepuestosText = Normalize-Text $repVtasSheet.Cells.Item($row, 24).Text
            CostoPintura = To-Number $repVtasSheet.Cells.Item($row, 25).Value2
            CostoPinturaText = Normalize-Text $repVtasSheet.Cells.Item($row, 25).Text
            CostoSubconNc = To-Number $repVtasSheet.Cells.Item($row, 26).Value2
            CostoSubconNcText = Normalize-Text $repVtasSheet.Cells.Item($row, 26).Text
            GarExt = Normalize-Text $repVtasSheet.Cells.Item($row, 27).Text
        }

        Add-LookupEntry -Store $repVtasStore -DocKey $docKey -OrderKey $orderKey -Entry $entry
    }

    return @{
        Invoice = $invoiceStore
        Note = $noteStore
        RepVtas = $repVtasStore
        Defaults = $defaults
    }
}

function Read-SourceRows {
    param([object]$Worksheet)

    Assert-NotCancelled 'lectura_fuente'
    $rows = New-Object System.Collections.Generic.List[object]
    $lastRow = $Worksheet.UsedRange.Row + $Worksheet.UsedRange.Rows.Count - 1

    for ($row = 2; $row -le $lastRow; $row++) {
        Assert-NotCancelled 'lectura_fuente'
        $agency = Normalize-Text $Worksheet.Cells.Item($row, 1).Text
        if ($agency -eq '') {
            continue
        }

        $order = Normalize-Text $Worksheet.Cells.Item($row, 3).Text
        $series = Normalize-Text $Worksheet.Cells.Item($row, 14).Text
        if ($series -eq '') {
            $series = Normalize-Text $Worksheet.Cells.Item($row, 13).Text
        }

        $templateKey = Get-Template-Key -Agency $agency -Series $series -Order $order
        if ($null -eq $templateKey) {
            continue
        }

        $docType = (Normalize-Text $Worksheet.Cells.Item($row, 8).Text).ToUpperInvariant()
        if ($docType -notin @('FA', 'FC', 'DC', 'DE')) {
            continue
        }

        $anulada = (Normalize-Text $Worksheet.Cells.Item($row, 36).Text).ToUpperInvariant()
        if ($anulada -in @('SI', 'S', 'YES', 'Y', 'ANULADA')) {
            continue
        }

        $documentRaw = Normalize-Text $Worksheet.Cells.Item($row, 12).Text
        if ($documentRaw -eq '') {
            continue
        }

        $affectedRaw = Normalize-Text $Worksheet.Cells.Item($row, 37).Text

        $rows.Add([pscustomobject]@{
            RowIndex = $row
            TemplateKey = $templateKey
            Agency = $agency
            AgencyRaw = [string]$Worksheet.Cells.Item($row, 1).Text
            Center = Normalize-Text $Worksheet.Cells.Item($row, 2).Text
            CenterRaw = [string]$Worksheet.Cells.Item($row, 2).Text
            Order = $order
            OrderRaw = [string]$Worksheet.Cells.Item($row, 3).Text
            Advisor = Normalize-Text $Worksheet.Cells.Item($row, 5).Text
            AdvisorRaw = [string]$Worksheet.Cells.Item($row, 5).Text
            Line = Normalize-Text $Worksheet.Cells.Item($row, 7).Text
            LineRaw = [string]$Worksheet.Cells.Item($row, 7).Text
            DocType = $docType
            Cedula = Normalize-Text $Worksheet.Cells.Item($row, 9).Text
            CedulaRaw = [string]$Worksheet.Cells.Item($row, 9).Text
            Customer = Normalize-Text $Worksheet.Cells.Item($row, 10).Text
            CustomerRaw = [string]$Worksheet.Cells.Item($row, 10).Text
            DocumentRaw = $documentRaw
            DocumentTrim = Trim-Document $documentRaw
            Series = $series
            SeriesRaw = [string]$series
            FormaPago = Normalize-Text $Worksheet.Cells.Item($row, 16).Text
            Authorization = Normalize-Text $Worksheet.Cells.Item($row, 17).Text
            DateFactValue = Get-Date-Value -Worksheet $Worksheet -Row $row -Column 15
            DateNoteValue = Get-Date-Value -Worksheet $Worksheet -Row $row -Column 18
            NoteCredit = To-Number $Worksheet.Cells.Item($row, 19).Value2
            TotalManoObra = To-Number $Worksheet.Cells.Item($row, 20).Value2
            TotalSubcontratos = To-Number $Worksheet.Cells.Item($row, 21).Value2
            TotalInsumos = To-Number $Worksheet.Cells.Item($row, 22).Value2
            TotalServicio = To-Number $Worksheet.Cells.Item($row, 23).Value2
            TotalAccesorios = To-Number $Worksheet.Cells.Item($row, 24).Value2
            TotalRepuestos = To-Number $Worksheet.Cells.Item($row, 25).Value2
            Interes = To-Number $Worksheet.Cells.Item($row, 26).Value2
            Iva = To-Number $Worksheet.Cells.Item($row, 27).Value2
            Total = To-Number $Worksheet.Cells.Item($row, 28).Value2
            Costo = To-Number $Worksheet.Cells.Item($row, 29).Value2
            CostoLubricantes = To-Number $Worksheet.Cells.Item($row, 30).Value2
            CostoAccesorios = To-Number $Worksheet.Cells.Item($row, 31).Value2
            CostoRepuestos = To-Number $Worksheet.Cells.Item($row, 32).Value2
            CostoPintura = To-Number $Worksheet.Cells.Item($row, 33).Value2
            CostoSubconNc = To-Number $Worksheet.Cells.Item($row, 34).Value2
            GarExt = Normalize-Text $Worksheet.Cells.Item($row, 35).Text
            GarExtRaw = [string]$Worksheet.Cells.Item($row, 35).Text
            Anulada = $anulada
            AffectedDocumentTrim = Trim-Document $affectedRaw
            AffectedDocumentRaw = [string]$Worksheet.Cells.Item($row, 37).Text
            MotivoNc = Normalize-Text $Worksheet.Cells.Item($row, 38).Text
            ObservacionNc = Normalize-Text $Worksheet.Cells.Item($row, 39).Text
        })
    }

    return $rows
}

function Normalize-SourceRows {
    param(
        [object[]]$Rows,
        [switch]$ConsolidateInvoiceDocuments
    )

    $sortedRows = @(
        $Rows |
            Sort-Object @{ Expression = { [int]$_.RowIndex } }
    )

    $normalized = New-Object System.Collections.Generic.List[object]
    $noteSeen = @{}
    $invoiceGroups = @{}

    foreach ($row in $sortedRows) {
        if ($row.DocType -in @('DC', 'DE')) {
            $dedupeKey = @(
                $row.TemplateKey,
                $row.DocType,
                $row.DocumentTrim,
                (Normalize-Text $row.Order),
                (Get-Date-Write-Value $row.DateNoteValue),
                (Round-Amount $row.NoteCredit),
                (Round-Amount $row.TotalManoObra),
                (Round-Amount $row.TotalSubcontratos),
                (Round-Amount $row.TotalInsumos),
                (Round-Amount $row.TotalServicio),
                (Round-Amount $row.TotalAccesorios),
                (Round-Amount $row.TotalRepuestos),
                (Round-Amount $row.Total),
                (Round-Amount $row.Iva),
                (Round-Amount $row.Interes),
                (Round-Amount $row.Anticipo),
                (Round-Amount $row.Neto),
                $row.AffectedDocumentTrim,
                $row.MotivoNc,
                $row.ObservacionNc
            ) -join '|'

            if ($noteSeen.ContainsKey($dedupeKey)) {
                continue
            }

            $noteSeen[$dedupeKey] = $true
            $normalized.Add($row)
            continue
        }

        if (-not $ConsolidateInvoiceDocuments) {
            $normalized.Add($row)
            continue
        }

        $invoiceGroupKey = @(
            $row.TemplateKey,
            $row.DocType,
            $row.DocumentTrim,
            (Normalize-Text $row.Series),
            (Get-Date-Write-Value $row.DateFactValue)
        ) -join '|'

        if (-not $invoiceGroups.ContainsKey($invoiceGroupKey)) {
            $invoiceGroups[$invoiceGroupKey] = New-Object System.Collections.Generic.List[object]
        }

        $invoiceGroups[$invoiceGroupKey].Add($row)
    }

    if ($ConsolidateInvoiceDocuments) {
        foreach ($group in $invoiceGroups.Values) {
            if ($group.Count -eq 1) {
                $normalized.Add($group[0])
                continue
            }

        $distinctOrders = @{}
        foreach ($item in $group) {
            $orderKey = Normalize-Text $item.Order
            if ($orderKey -ne '') {
                $distinctOrders[$orderKey] = $true
            }
        }

        if ($distinctOrders.Count -le 1) {
            # Duplicado exacto de factura: conservar solo una fila.
            $normalized.Add($group[0])
            continue
        }

        # Misma factura repetida en varias ordenes (caso operativo SZK):
        # consolidar montos en una sola fila con orden en blanco.
        $first = @($group | Sort-Object @{ Expression = { [int]$_.RowIndex } })[0]
        $sumFields = @(
            'NoteCredit',
            'TotalManoObra',
            'TotalSubcontratos',
            'TotalInsumos',
            'TotalServicio',
            'TotalAccesorios',
            'TotalRepuestos',
            'Interes',
            'Iva12',
            'Iva15',
            'Iva',
            'Total',
            'Costo',
            'CostoLubricantes',
            'CostoAccesorios',
            'CostoRepuestos',
            'CostoPintura',
            'CostoSubconNc'
        )

        $sumValues = @{}
        foreach ($field in $sumFields) {
            $sumValues[$field] = [double]0
        }

        foreach ($item in $group) {
            foreach ($field in $sumFields) {
                $sumValues[$field] = [double]$sumValues[$field] + [double](To-Number $item.$field)
            }
        }

            $normalized.Add([pscustomobject]@{
                RowIndex = $first.RowIndex
                TemplateKey = $first.TemplateKey
                Agency = $first.Agency
                AgencyRaw = $first.AgencyRaw
                Center = $first.Center
                CenterRaw = $first.CenterRaw
                Order = ''
                OrderRaw = ''
                Advisor = $first.Advisor
                AdvisorRaw = $first.AdvisorRaw
                Line = $first.Line
                LineRaw = $first.LineRaw
                DocType = $first.DocType
                Cedula = $first.Cedula
                CedulaRaw = $first.CedulaRaw
                Customer = $first.Customer
                CustomerRaw = $first.CustomerRaw
                DocumentRaw = $first.DocumentRaw
                DocumentTrim = $first.DocumentTrim
                Series = $first.Series
                SeriesRaw = $first.SeriesRaw
                FormaPago = $first.FormaPago
                Authorization = $first.Authorization
                DateFactValue = $first.DateFactValue
                DateNoteValue = $first.DateNoteValue
                NoteCredit = $sumValues['NoteCredit']
                TotalManoObra = $sumValues['TotalManoObra']
                TotalSubcontratos = $sumValues['TotalSubcontratos']
                TotalInsumos = $sumValues['TotalInsumos']
                TotalServicio = $sumValues['TotalServicio']
                TotalAccesorios = $sumValues['TotalAccesorios']
                TotalRepuestos = $sumValues['TotalRepuestos']
                Interes = $sumValues['Interes']
                Iva12 = $sumValues['Iva12']
                Iva15 = $sumValues['Iva15']
                Iva = $sumValues['Iva']
                Total = $sumValues['Total']
                Costo = $sumValues['Costo']
                CostoLubricantes = $sumValues['CostoLubricantes']
                CostoAccesorios = $sumValues['CostoAccesorios']
                CostoRepuestos = $sumValues['CostoRepuestos']
                CostoPintura = $sumValues['CostoPintura']
                CostoSubconNc = $sumValues['CostoSubconNc']
                GarExt = $first.GarExt
                GarExtRaw = $first.GarExtRaw
                Anulada = $first.Anulada
                AffectedDocumentTrim = $first.AffectedDocumentTrim
                AffectedDocumentRaw = $first.AffectedDocumentRaw
                MotivoNc = $first.MotivoNc
                ObservacionNc = $first.ObservacionNc
            })
        }
    }

    return @(
        $normalized |
            Sort-Object `
                @{ Expression = { [int]$_.RowIndex } }, `
                @{ Expression = { Normalize-Text $_.TemplateKey } }, `
                @{ Expression = { Normalize-Text $_.DocType } }, `
                @{ Expression = { Get-Document-SortValue $_.DocumentTrim } }, `
                @{ Expression = { Normalize-Text $_.Order } }
    )
}

function Clear-OutputSheet {
    param(
        [object]$Worksheet,
        [int]$StartRow,
        [string]$LastColumn
    )

    $range = $Worksheet.Range("A${StartRow}:${LastColumn}65536")
    $xlCellTypeConstants = 2
    try {
        $constants = $range.SpecialCells($xlCellTypeConstants)
        $null = $constants.ClearContents()
        [void][Runtime.Interopservices.Marshal]::ReleaseComObject($constants)
    } catch {
        # Si no hay constantes en el rango, no hay nada que limpiar.
    }
    [void][Runtime.Interopservices.Marshal]::ReleaseComObject($range)
}

function Clear-Worksheet-UsedRange {
    param([object]$Worksheet)

    try {
        $Worksheet.UsedRange.ClearContents() | Out-Null
    } catch {
    }
}

function Clear-Worksheet-RangeContents {
    param(
        [object]$Worksheet,
        [string]$StartColumn = 'A',
        [int]$StartRow = 1,
        [string]$EndColumn = 'A',
        [int]$EndRow = 1
    )

    if ($null -eq $Worksheet -or $StartRow -gt $EndRow) {
        return
    }

    $range = $null
    try {
        $range = $Worksheet.Range("${StartColumn}${StartRow}:${EndColumn}${EndRow}")
        $null = $range.ClearContents()
    } finally {
        if ($null -ne $range) {
            [void][Runtime.Interopservices.Marshal]::ReleaseComObject($range)
        }
    }
}

function Write-Rows-ToWorksheet {
    param(
        [object]$Worksheet,
        [object[]]$Rows,
        [int]$StartRow = 1,
        [int]$StartColumn = 1
    )

    if ($null -eq $Rows -or $Rows.Count -eq 0) {
        return
    }

    $maxColumns = 0
    foreach ($row in $Rows) {
        $count = if ($row -is [System.Collections.ICollection]) { $row.Count } else { 0 }
        if ($count -gt $maxColumns) {
            $maxColumns = $count
        }
    }
    if ($maxColumns -le 0) { return }

    $rowCount = [int]$Rows.Count
    $colCount = [int]$maxColumns
    $matrix = [System.Array]::CreateInstance([object], @($rowCount, $colCount), @(1, 1))
    for ($r = 0; $r -lt $rowCount; $r++) {
        $row = $Rows[$r]
        for ($c = 0; $c -lt $colCount; $c++) {
            $value = $null
            if ($row -is [System.Collections.ICollection] -and $c -lt $row.Count) {
                $value = $row[$c]
            }

            $matrix.SetValue($value, $r + 1, $c + 1)
        }
    }

    $startCell = $null
    $endCell = $null
    $range = $null
    try {
        $startCell = $Worksheet.Cells.Item($StartRow, $StartColumn)
        $endCell = $Worksheet.Cells.Item($StartRow + $rowCount - 1, $StartColumn + $colCount - 1)
        $range = $Worksheet.Range($startCell, $endCell)
        $range.Value2 = $matrix
    } finally {
        if ($null -ne $range) {
            [void][Runtime.Interopservices.Marshal]::ReleaseComObject($range)
        }
        if ($null -ne $endCell) {
            [void][Runtime.Interopservices.Marshal]::ReleaseComObject($endCell)
        }
        if ($null -ne $startCell) {
            [void][Runtime.Interopservices.Marshal]::ReleaseComObject($startCell)
        }
    }
}

function Get-DateColumnWriteMode {
    param(
        [object]$Worksheet,
        [int]$Row,
        [int]$Column
    )

    $cell = $null
    try {
        $cell = $Worksheet.Cells.Item($Row, $Column)
        $format = Normalize-Text $cell.NumberFormat
        if ($format -eq '' -or $format -eq 'General') {
            return 'text'
        }

        return 'serial'
    } finally {
        if ($null -ne $cell) {
            [void][Runtime.Interopservices.Marshal]::ReleaseComObject($cell)
        }
    }
}

function Get-DateMatrixValue {
    param(
        [object]$Value,
        [string]$Mode = 'serial',
        [string]$TextFormat = 'dd/MM/yyyy'
    )

    if ($null -eq $Value -or $Value -eq '') {
        return $null
    }

    $dateSerial = [double]$Value
    if ($Mode -eq 'text') {
        return "'" + [datetime]::FromOADate($dateSerial).ToString($TextFormat)
    }

    return $dateSerial
}

function Get-NumericMatrixValue {
    param(
        [object]$Value,
        [switch]$BlankIfZero
    )

    $rounded = [double](Round-Amount (To-Number $Value))
    if ($BlankIfZero -and (Round-Amount ([Math]::Abs($rounded))) -eq 0) {
        return $null
    }

    return $rounded
}

function Resolve-MayorPathForBrand {
    param(
        [string]$BrandKey,
        [hashtable]$MayorPaths
    )

    $key = (Normalize-Text $BrandKey).ToLowerInvariant()
    if ($MayorPaths.ContainsKey($key)) {
        $specificPath = Normalize-Text $MayorPaths[$key]
        if ($specificPath -ne '') {
            return $specificPath
        }
    }

    return ''
}

function Get-MayorSheetSectionLayouts {
    param([object]$Worksheet)

    $layouts = New-Object System.Collections.Generic.List[object]
    $summaryRanges = New-Object System.Collections.Generic.List[object]
    $usedRange = $null
    $lastRow = 0
    try {
        $usedRange = $Worksheet.UsedRange
        $lastRow = [int]($usedRange.Row + $usedRange.Rows.Count - 1)
    } finally {
        if ($null -ne $usedRange) {
            [void][Runtime.Interopservices.Marshal]::ReleaseComObject($usedRange)
        }
    }

    for ($row = 1; $row -le $lastRow; $row++) {
        $formulaText = Normalize-Text $Worksheet.Cells.Item($row, 9).Formula
        if ($formulaText -eq '') {
            continue
        }

        $normalizedFormula = $formulaText.ToUpperInvariant().Replace('$', '')
        if ($normalizedFormula -notmatch 'SUBTOTAL\(9,I(\d+):I(\d+)\)') {
            continue
        }

        $summaryRanges.Add([pscustomobject]@{
            StartRow = [int]$Matches[1]
            EndRow = [int]$Matches[2]
        }) | Out-Null
    }

    if ($summaryRanges.Count -eq 0) {
        return @()
    }

    $accountStarts = New-Object System.Collections.Generic.List[object]
    $seenAccounts = @{}
    for ($row = 1; $row -le $lastRow; $row++) {
        $account = Normalize-Text $Worksheet.Cells.Item($row, 1).Text
        if ($account -notmatch '^\d{2}\.\d{2}\.\d{2}\.\d{2}\.\d{4}$' -or $seenAccounts.ContainsKey($account)) {
            continue
        }

        $parentRange = $null
        foreach ($range in $summaryRanges) {
            if ($row -ge [int]$range.StartRow -and $row -le [int]$range.EndRow) {
                $parentRange = $range
                break
            }
        }

        if ($null -eq $parentRange) {
            continue
        }

        $seenAccounts[$account] = $true
        $accountStarts.Add([pscustomobject]@{
            Account = $account
            Name = Normalize-Text $Worksheet.Cells.Item($row, 2).Text
            StartRow = [int]$row
            ParentStartRow = [int]$parentRange.StartRow
            ParentEndRow = [int]$parentRange.EndRow
        }) | Out-Null
    }

    $orderedStarts = @($accountStarts | Sort-Object @{ Expression = { [int]$_.StartRow } })
    for ($index = 0; $index -lt $orderedStarts.Count; $index++) {
        $current = $orderedStarts[$index]
        $endRow = [int]$current.ParentEndRow
        for ($nextIndex = $index + 1; $nextIndex -lt $orderedStarts.Count; $nextIndex++) {
            $next = $orderedStarts[$nextIndex]
            if ([int]$next.ParentStartRow -ne [int]$current.ParentStartRow) {
                break
            }

            $endRow = [int]$next.StartRow - 1
            break
        }

        $layouts.Add([pscustomobject]@{
            Account = $current.Account
            Name = $current.Name
            StartRow = [int]$current.StartRow
            EndRow = [int]$endRow
            ParentStartRow = [int]$current.ParentStartRow
            ParentEndRow = [int]$current.ParentEndRow
        }) | Out-Null
    }

    return $layouts.ToArray()
}

function Get-MayorAccountFamilyKey {
    param([string]$Account)

    $normalized = Normalize-Text $Account
    if ($normalized -notmatch '^(\d{2}\.\d{2}\.\d{2}\.\d{2})\.(\d{4})$') {
        return $normalized
    }

    $prefix = $Matches[1]
    $suffix = $Matches[2]
    switch ($suffix) {
        { $_ -in @('0001', '0002', '0003') } { return "$prefix|VENTA" }
        { $_ -in @('0010', '0012') } { return "$prefix|DESCUENTO" }
        '0014' { return "$prefix|DEVOLUCION" }
        default { return "$prefix|$suffix" }
    }
}

function Resolve-MayorCompatibleLayout {
    param(
        [string]$Account,
        [object[]]$Layouts,
        [hashtable]$RowsByLayoutKey
    )

    $familyKey = Get-MayorAccountFamilyKey -Account $Account
    if ($familyKey -eq '') {
        return $null
    }

    $candidates = @(
        @($Layouts) |
            Where-Object { (Get-MayorAccountFamilyKey -Account $_.Account) -eq $familyKey } |
            Sort-Object @{ Expression = { [int]$_.EndRow - [int]$_.StartRow + 1 }; Descending = $true }, @{ Expression = { [int]$_.StartRow } }
    )
    if ($candidates.Count -eq 0) {
        return $null
    }

    foreach ($candidate in $candidates) {
        $layoutKey = "{0}:{1}-{2}" -f $candidate.Account, [int]$candidate.StartRow, [int]$candidate.EndRow
        $existing = 0
        if ($RowsByLayoutKey.ContainsKey($layoutKey)) {
            $existing = [int]$RowsByLayoutKey[$layoutKey].Count
        }

        $capacity = [int]($candidate.EndRow - $candidate.StartRow + 1)
        if ($existing -lt $capacity) {
            return $candidate
        }
    }

    return $candidates[0]
}

function Write-MayorDataRow {
    param(
        [object]$Worksheet,
        [int]$TargetRow,
        [pscustomobject]$MayorRow
    )

    $null = $Worksheet.Cells.Item($TargetRow, 1).Value2 = (Get-Excel-TextLiteral (Normalize-Text $MayorRow.account))
    $null = $Worksheet.Cells.Item($TargetRow, 2).Value2 = (Get-Excel-TextLiteral (Normalize-Text $MayorRow.name))
    $null = $Worksheet.Cells.Item($TargetRow, 3).Value2 = (Get-Excel-TextLiteral (Normalize-Text $MayorRow.ext))
    Set-DateCellValue -Worksheet $Worksheet -Row $TargetRow -Column 4 -Value $MayorRow.date_value -Context 'MAYOR_FECHA'
    $null = $Worksheet.Cells.Item($TargetRow, 5).Value2 = (Get-Excel-TextLiteral (Normalize-Text $MayorRow.origin))
    $null = $Worksheet.Cells.Item($TargetRow, 6).Value2 = (Get-Excel-TextLiteral (Normalize-Text $MayorRow.seat))
    $null = $Worksheet.Cells.Item($TargetRow, 7).Value2 = (Get-Excel-TextLiteral (Normalize-Text $MayorRow.reference))
    $null = $Worksheet.Cells.Item($TargetRow, 8).Value2 = (Get-Excel-TextLiteral (Normalize-Text $MayorRow.detail))
    Set-NumericCellSafe -Worksheet $Worksheet -Row $TargetRow -Column 9 -Value (To-Number $MayorRow.debit)
    Set-NumericCellSafe -Worksheet $Worksheet -Row $TargetRow -Column 10 -Value (To-Number $MayorRow.credit)
    $balanceValue = if ($null -ne $MayorRow.PSObject.Properties['effective_balance']) { $MayorRow.effective_balance } else { $MayorRow.balance }
    Set-NumericCellSafe -Worksheet $Worksheet -Row $TargetRow -Column 11 -Value (To-Number $balanceValue)
}

function Write-MayorRows-ToWorksheet {
    param(
        [object]$Worksheet,
        [object[]]$Rows,
        [string]$BrandKey = ''
    )

    $layouts = @(Get-MayorSheetSectionLayouts -Worksheet $Worksheet)
    if ($layouts.Count -eq 0) {
        throw "La hoja $($Worksheet.Name) no tiene secciones SUBTOTAL reconocibles para cargar el mayor."
    }

    $rowsByAccount = @{}
    foreach ($row in @($Rows)) {
        $account = Normalize-Text $row.account
        if ($account -eq '') {
            continue
        }

        if (-not $rowsByAccount.ContainsKey($account)) {
            $rowsByAccount[$account] = New-Object System.Collections.Generic.List[object]
        }

        $rowsByAccount[$account].Add($row) | Out-Null
    }

    $knownAccounts = @{}
    $rowsByLayoutKey = @{}
    $compatibleMappedAccounts = New-Object System.Collections.Generic.List[string]
    $unmappedAccounts = New-Object System.Collections.Generic.List[string]

    foreach ($account in @($rowsByAccount.Keys)) {
        $layout = @($layouts | Where-Object { $_.Account -eq $account } | Select-Object -First 1)[0]
        if ($null -eq $layout) {
            $layout = Resolve-MayorCompatibleLayout -Account $account -Layouts $layouts -RowsByLayoutKey $rowsByLayoutKey
            if ($null -eq $layout) {
                Write-Output ("WARN|mayor_account_unmapped|{0}|{1}|sheet={2}" -f $BrandKey, $account, $Worksheet.Name)
                $unmappedAccounts.Add([string]$account) | Out-Null
                continue
            }

            $compatibleMappedAccounts.Add([string]$account) | Out-Null
            Write-Output ("WARN|mayor_account_compatible_section|{0}|{1}|section={2}|sheet={3}" -f $BrandKey, $account, $layout.Account, $Worksheet.Name)
        }

        $layoutKey = "{0}:{1}-{2}" -f $layout.Account, [int]$layout.StartRow, [int]$layout.EndRow
        if (-not $rowsByLayoutKey.ContainsKey($layoutKey)) {
            $rowsByLayoutKey[$layoutKey] = New-Object System.Collections.Generic.List[object]
        }

        foreach ($accountRow in $rowsByAccount[$account].ToArray()) {
            $rowsByLayoutKey[$layoutKey].Add($accountRow) | Out-Null
        }
    }

    $writtenCount = 0
    foreach ($layout in $layouts) {
        $knownAccounts[$layout.Account] = $true
        Clear-Worksheet-RangeContents -Worksheet $Worksheet -StartColumn 'A' -StartRow $layout.StartRow -EndColumn 'M' -EndRow $layout.EndRow
        $layoutKey = "{0}:{1}-{2}" -f $layout.Account, [int]$layout.StartRow, [int]$layout.EndRow
        $accountRows = @()
        if ($rowsByLayoutKey.ContainsKey($layoutKey)) {
            $accountRows = $rowsByLayoutKey[$layoutKey].ToArray()
        }
        $capacity = [int]($layout.EndRow - $layout.StartRow + 1)
        if ($accountRows.Count -gt $capacity) {
            throw ("El mayor para {0} excede la capacidad de la seccion {1} en {2}. Capacidad={3}, filas={4}" -f $BrandKey, $layout.Account, $Worksheet.Name, $capacity, $accountRows.Count)
        }

        if ($accountRows.Count -gt 0) {
            $dateMode = Get-DateColumnWriteMode -Worksheet $Worksheet -Row ([int]$layout.StartRow) -Column 4
            $sheetRows = New-Object System.Collections.Generic.List[object]
            foreach ($accountRow in $accountRows) {
                $balanceValue = if ($null -ne $accountRow.PSObject.Properties['effective_balance']) { $accountRow.effective_balance } else { $accountRow.balance }
                $sheetRows.Add(@(
                    (Get-Excel-TextLiteral (Normalize-Text $accountRow.account)),
                    (Get-Excel-TextLiteral (Normalize-Text $accountRow.name)),
                    (Get-Excel-TextLiteral (Normalize-Text $accountRow.ext)),
                    (Get-DateMatrixValue -Value $accountRow.date_value -Mode $dateMode),
                    (Get-Excel-TextLiteral (Normalize-Text $accountRow.origin)),
                    (Get-Excel-TextLiteral (Normalize-Text $accountRow.seat)),
                    (Get-Excel-TextLiteral (Normalize-Text $accountRow.reference)),
                    (Get-Excel-TextLiteral (Normalize-Text $accountRow.detail)),
                    (Get-NumericMatrixValue -Value $accountRow.debit),
                    (Get-NumericMatrixValue -Value $accountRow.credit),
                    (Get-NumericMatrixValue -Value $balanceValue)
                )) | Out-Null
            }

            Write-Rows-ToWorksheet -Worksheet $Worksheet -Rows $sheetRows.ToArray() -StartRow ([int]$layout.StartRow) -StartColumn 1
            $writtenCount += $accountRows.Count
        }
    }

    $usedRange = $null
    try {
        $usedRange = $Worksheet.UsedRange
        $lastRow = [int]($usedRange.Row + $usedRange.Rows.Count - 1)
        for ($row = 1; $row -le $lastRow; $row++) {
            $accountText = Normalize-Text $Worksheet.Cells.Item($row, 1).Text
            if ($accountText -notmatch '^\d{2}\.\d{2}\.\d{2}\.\d{2}\.\d{4}$') {
                continue
            }

            $insideLayout = $false
            foreach ($layout in $layouts) {
                if ($row -ge [int]$layout.StartRow -and $row -le [int]$layout.EndRow) {
                    $insideLayout = $true
                    break
                }
            }

            if (-not $insideLayout) {
                Clear-Worksheet-RangeContents -Worksheet $Worksheet -StartColumn 'A' -StartRow $row -EndColumn 'M' -EndRow $row
            }
        }
    } finally {
        if ($null -ne $usedRange) {
            [void][Runtime.Interopservices.Marshal]::ReleaseComObject($usedRange)
        }
    }

    return [pscustomobject]@{
        RowCount = $writtenCount
        SectionCount = $layouts.Count
        UnmappedAccounts = $unmappedAccounts.ToArray()
        CompatibleMappedAccounts = $compatibleMappedAccounts.ToArray()
    }
}

function Copy-WorksheetPageSetupFromTemplate {
    param(
        [object]$TemplateWorksheet,
        [object]$OutputWorksheet
    )

    if ($null -eq $TemplateWorksheet -or $null -eq $OutputWorksheet) {
        return
    }

    $templatePageSetup = $null
    $outputPageSetup = $null
    try {
        $templatePageSetup = $TemplateWorksheet.PageSetup
        $outputPageSetup = $OutputWorksheet.PageSetup
        foreach ($propertyName in @(
            'PrintArea',
            'Zoom',
            'FitToPagesWide',
            'FitToPagesTall',
            'Orientation',
            'PaperSize',
            'LeftMargin',
            'RightMargin',
            'TopMargin',
            'BottomMargin',
            'HeaderMargin',
            'FooterMargin',
            'CenterHorizontally',
            'CenterVertically',
            'PrintGridlines',
            'PrintHeadings',
            'BlackAndWhite',
            'Draft',
            'PrintComments',
            'Order'
        )) {
            try {
                $outputPageSetup.$propertyName = $templatePageSetup.$propertyName
            } catch {}
        }
    } finally {
        if ($null -ne $outputPageSetup) {
            [void][Runtime.Interopservices.Marshal]::ReleaseComObject($outputPageSetup)
        }
        if ($null -ne $templatePageSetup) {
            [void][Runtime.Interopservices.Marshal]::ReleaseComObject($templatePageSetup)
        }
    }
}

function Get-WorksheetMeaningfulBounds {
    param([object]$Worksheet)

    if ($null -eq $Worksheet) {
        return [pscustomobject]@{
            LastRow = 1
            LastColumn = 1
        }
    }

    $afterCell = $null
    $lastRowCell = $null
    $lastColumnCell = $null
    try {
        $afterCell = $Worksheet.Cells.Item(1, 1)
        $lastRowCell = $Worksheet.Cells.Find('*', $afterCell, -4123, 2, 1, 2, $false, $false, $false)
        $lastColumnCell = $Worksheet.Cells.Find('*', $afterCell, -4123, 2, 2, 2, $false, $false, $false)

        return [pscustomobject]@{
            LastRow = if ($null -ne $lastRowCell) { [int]$lastRowCell.Row } else { 1 }
            LastColumn = if ($null -ne $lastColumnCell) { [int]$lastColumnCell.Column } else { 1 }
        }
    } catch {
        $usedRange = $null
        try {
            $usedRange = $Worksheet.UsedRange
            return [pscustomobject]@{
                LastRow = [Math]::Max(1, [int]($usedRange.Row + $usedRange.Rows.Count - 1))
                LastColumn = [Math]::Max(1, [int]($usedRange.Column + $usedRange.Columns.Count - 1))
            }
        } finally {
            if ($null -ne $usedRange) {
                [void][Runtime.Interopservices.Marshal]::ReleaseComObject($usedRange)
            }
        }
    } finally {
        if ($null -ne $lastColumnCell) {
            [void][Runtime.Interopservices.Marshal]::ReleaseComObject($lastColumnCell)
        }
        if ($null -ne $lastRowCell) {
            [void][Runtime.Interopservices.Marshal]::ReleaseComObject($lastRowCell)
        }
        if ($null -ne $afterCell) {
            [void][Runtime.Interopservices.Marshal]::ReleaseComObject($afterCell)
        }
    }
}

function Restore-WorksheetLayoutFromTemplate {
    param(
        [object]$TemplateWorksheet,
        [object]$OutputWorksheet,
        [int]$StartRow = 1,
        [int]$EndRow = 0,
        [int]$LastColumn = 0
    )

    if ($null -eq $TemplateWorksheet -or $null -eq $OutputWorksheet) {
        return
    }

    $bounds = Get-WorksheetMeaningfulBounds -Worksheet $TemplateWorksheet
    $templateLastRow = [Math]::Max(1, [int]$bounds.LastRow)
    $templateLastColumn = [Math]::Max(1, [int]$bounds.LastColumn)
    $effectiveStartRow = [Math]::Max(1, [int]$StartRow)
    $effectiveEndRow = if ($EndRow -gt 0) { [Math]::Min($templateLastRow, [int]$EndRow) } else { $templateLastRow }
    $effectiveLastColumn = if ($LastColumn -gt 0) { [Math]::Min($templateLastColumn, [int]$LastColumn) } else { $templateLastColumn }

    for ($column = 1; $column -le $effectiveLastColumn; $column++) {
        $templateColumn = $null
        $outputColumn = $null
        try {
            $templateColumn = $TemplateWorksheet.Columns.Item($column)
            $outputColumn = $OutputWorksheet.Columns.Item($column)
            try { $outputColumn.ColumnWidth = $templateColumn.ColumnWidth } catch {}
            try { $outputColumn.Hidden = $templateColumn.Hidden } catch {}
        } finally {
            if ($null -ne $outputColumn) {
                [void][Runtime.Interopservices.Marshal]::ReleaseComObject($outputColumn)
            }
            if ($null -ne $templateColumn) {
                [void][Runtime.Interopservices.Marshal]::ReleaseComObject($templateColumn)
            }
        }
    }

    for ($row = $effectiveStartRow; $row -le $effectiveEndRow; $row++) {
        $templateRow = $null
        $outputRow = $null
        try {
            $templateRow = $TemplateWorksheet.Rows.Item($row)
            $outputRow = $OutputWorksheet.Rows.Item($row)
            try { $outputRow.RowHeight = $templateRow.RowHeight } catch {}
            try { $outputRow.Hidden = $templateRow.Hidden } catch {}
        } finally {
            if ($null -ne $outputRow) {
                [void][Runtime.Interopservices.Marshal]::ReleaseComObject($outputRow)
            }
            if ($null -ne $templateRow) {
                [void][Runtime.Interopservices.Marshal]::ReleaseComObject($templateRow)
            }
        }
    }

    try { $OutputWorksheet.StandardWidth = $TemplateWorksheet.StandardWidth } catch {}
    try { $OutputWorksheet.DefaultRowHeight = $TemplateWorksheet.DefaultRowHeight } catch {}
    Copy-WorksheetPageSetupFromTemplate -TemplateWorksheet $TemplateWorksheet -OutputWorksheet $OutputWorksheet
}

function Restore-WorkbookLayoutFromTemplate {
    param(
        [object]$TemplateWorkbook,
        [object]$OutputWorkbook
    )

    if ($null -eq $TemplateWorkbook -or $null -eq $OutputWorkbook) {
        return
    }

    foreach ($templateSheet in @($TemplateWorkbook.Worksheets)) {
        $outputSheet = $null
        try {
            try { $outputSheet = $OutputWorkbook.Worksheets.Item($templateSheet.Name) } catch {}
            if ($null -ne $outputSheet) {
                Restore-WorksheetLayoutFromTemplate -TemplateWorksheet $templateSheet -OutputWorksheet $outputSheet
            }
        } finally {
            if ($null -ne $outputSheet) {
                [void][Runtime.Interopservices.Marshal]::ReleaseComObject($outputSheet)
            }
            if ($null -ne $templateSheet) {
                [void][Runtime.Interopservices.Marshal]::ReleaseComObject($templateSheet)
            }
        }
    }
}

function Restore-SelectedWorkbookLayoutFromTemplate {
    param(
        [object]$TemplateWorkbook,
        [object]$OutputWorkbook,
        [string[]]$SheetNames,
        [hashtable]$SheetRangeMap = @{}
    )

    if ($null -eq $TemplateWorkbook -or $null -eq $OutputWorkbook) {
        return
    }

    $normalizedNames = @(
        @($SheetNames) |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            ForEach-Object { Normalize-Text $_ } |
            Select-Object -Unique
    )
    if ($normalizedNames.Count -eq 0) {
        Restore-WorkbookLayoutFromTemplate -TemplateWorkbook $TemplateWorkbook -OutputWorkbook $OutputWorkbook
        return
    }

    foreach ($sheetName in $normalizedNames) {
        $templateSheet = $null
        $outputSheet = $null
        try {
            try { $templateSheet = $TemplateWorkbook.Worksheets.Item($sheetName) } catch {}
            try { $outputSheet = $OutputWorkbook.Worksheets.Item($sheetName) } catch {}
            if ($null -ne $templateSheet -and $null -ne $outputSheet) {
                $sheetRange = $null
                if ($SheetRangeMap.ContainsKey($sheetName)) {
                    $sheetRange = $SheetRangeMap[$sheetName]
                }

                if ($null -ne $sheetRange) {
                    Restore-WorksheetLayoutFromTemplate `
                        -TemplateWorksheet $templateSheet `
                        -OutputWorksheet $outputSheet `
                        -StartRow $(if ($null -ne $sheetRange.StartRow) { [int]$sheetRange.StartRow } else { 1 }) `
                        -EndRow $(if ($null -ne $sheetRange.EndRow) { [int]$sheetRange.EndRow } else { 0 }) `
                        -LastColumn $(if ($null -ne $sheetRange.LastColumn) { [int]$sheetRange.LastColumn } else { 0 })
                } else {
                    Restore-WorksheetLayoutFromTemplate -TemplateWorksheet $templateSheet -OutputWorksheet $outputSheet
                }
            }
        } finally {
            if ($null -ne $outputSheet) {
                [void][Runtime.Interopservices.Marshal]::ReleaseComObject($outputSheet)
            }
            if ($null -ne $templateSheet) {
                [void][Runtime.Interopservices.Marshal]::ReleaseComObject($templateSheet)
            }
        }
    }
}

function Get-CompactAccountCode {
    param([object]$Value)

    $text = Normalize-Text $Value
    if ($text -eq '') {
        return ''
    }

    $digits = ($text -replace '[^0-9]', '')
    if ($digits.Length -gt 0 -and $digits.Length -lt 12) {
        return $digits.PadLeft(12, '0')
    }
    return $digits
}

function Get-NormalizedCenterCode {
    param([object]$Value)

    $text = Normalize-Text $Value
    if ($text -eq '') {
        return ''
    }

    if ($text -match '^\d+$') {
        return $text.PadLeft(2, '0')
    }

    return $text
}

function Get-Brand-PeriodDateValue {
    param(
        [object[]]$Rows,
        [object[]]$MayorRows
    )

    $maxDate = $null

    foreach ($row in @($Rows)) {
        foreach ($candidate in @($row.DateFactValue, $row.DateNoteValue)) {
            if ($null -eq $candidate -or $candidate -eq '') {
                continue
            }

            $candidateValue = [double]$candidate
            if ($null -eq $maxDate -or $candidateValue -gt [double]$maxDate) {
                $maxDate = $candidateValue
            }
        }
    }

    foreach ($row in @($MayorRows)) {
        $candidate = $row.date_value
        if ($null -eq $candidate -or $candidate -eq '') {
            continue
        }

        $candidateValue = [double]$candidate
        if ($null -eq $maxDate -or $candidateValue -gt [double]$maxDate) {
            $maxDate = $candidateValue
        }
    }

    if ($null -eq $maxDate) {
        return (Get-Date).ToOADate()
    }

    return [double]$maxDate
}

function Add-GeneratedTemplateRow {
    param(
        [System.Collections.Generic.List[object]]$Target,
        [pscustomobject]$Prototype,
        [double]$Debit = 0.0,
        [double]$Credit = 0.0
    )

    if ($null -eq $Prototype) {
        return
    }

    $Target.Add([pscustomobject]@{
        Ag = Normalize-Text $Prototype.Ag
        Doc = Normalize-Text $Prototype.Doc
        Line = Normalize-Text $Prototype.Line
        Account = Normalize-Text $Prototype.Account
        Description = Normalize-Text $Prototype.Description
        CostCenter = Normalize-Text $Prototype.CostCenter
        Debit = [double](Round-Amount $Debit)
        Credit = [double](Round-Amount $Credit)
        Asiento = Normalize-Text $Prototype.Asiento
    }) | Out-Null
}

function Add-TemplateAggregateRows {
    param(
        [System.Collections.Generic.List[object]]$Target,
        [object[]]$Prototypes,
        [hashtable]$TotalsByCenter,
        [string]$AmountSide = 'Debit',
        [switch]$EnsureAtLeastOneRow
    )

    $orderedPrototypes = @($Prototypes)
    if ($orderedPrototypes.Count -eq 0) {
        return
    }

    $written = $false
    foreach ($center in @($TotalsByCenter.Keys | Sort-Object)) {
        $prototype = @($orderedPrototypes | Where-Object { (Normalize-Text $_.Ag) -eq $center } | Select-Object -First 1)[0]
        if ($null -eq $prototype) {
            $prototype = $orderedPrototypes[0]
        }

        $amount = [double](Round-Amount ([double]$TotalsByCenter[$center]))
        if ($AmountSide -eq 'Credit') {
            Add-GeneratedTemplateRow -Target $Target -Prototype $prototype -Debit 0.0 -Credit $amount
        } else {
            Add-GeneratedTemplateRow -Target $Target -Prototype $prototype -Debit $amount -Credit 0.0
        }

        $written = $true
    }

    if (-not $written -and $EnsureAtLeastOneRow) {
        if ($AmountSide -eq 'Credit') {
            Add-GeneratedTemplateRow -Target $Target -Prototype $orderedPrototypes[0] -Debit 0.0 -Credit 0.0
        } else {
            Add-GeneratedTemplateRow -Target $Target -Prototype $orderedPrototypes[0] -Debit 0.0 -Credit 0.0
        }
    }
}

function Add-TemplateSequentialRowsFromMayor {
    param(
        [System.Collections.Generic.List[object]]$Target,
        [object[]]$Prototypes,
        [object[]]$MayorRows,
        [switch]$EnsureAtLeastOneRow
    )

    $orderedPrototypes = @($Prototypes)
    if ($orderedPrototypes.Count -eq 0) {
        return
    }

    $items = @($MayorRows)
    if ($items.Count -eq 0 -and $EnsureAtLeastOneRow) {
        Add-GeneratedTemplateRow -Target $Target -Prototype $orderedPrototypes[0] -Debit 0.0 -Credit 0.0
        return
    }

    for ($index = 0; $index -lt $items.Count; $index++) {
        $prototype = if ($index -lt $orderedPrototypes.Count) { $orderedPrototypes[$index] } else { $orderedPrototypes[$orderedPrototypes.Count - 1] }
        $item = $items[$index]
        Add-GeneratedTemplateRow -Target $Target -Prototype $prototype -Debit (To-Number $item.debit) -Credit (To-Number $item.credit)
    }
}

function Read-PrecontVentasPrototypes {
    param([object]$Worksheet)

    $prototypes = New-Object System.Collections.Generic.List[object]
    $usedRange = $null
    try {
        $usedRange = $Worksheet.UsedRange
        $lastRow = [int]($usedRange.Row + $usedRange.Rows.Count - 1)
        for ($row = 2; $row -le $lastRow; $row++) {
            $accountText = Normalize-Text $Worksheet.Cells.Item($row, 5).Text
            $accountDigits = ($accountText -replace '[^0-9]', '')
            $account = if ($accountDigits.Length -ge 10) { $accountDigits.PadLeft(12, '0') } else { $accountText }
            $doc = Normalize-Text $Worksheet.Cells.Item($row, 3).Text
            if ($account -eq '' -or $doc -eq '') {
                continue
            }

            $prototypes.Add([pscustomobject]@{
                TemplateRow = $row
                Ag = Normalize-Text $Worksheet.Cells.Item($row, 2).Text
                Doc = $doc
                Line = Normalize-Text $Worksheet.Cells.Item($row, 4).Text
                Account = $account
                Description = Normalize-Text $Worksheet.Cells.Item($row, 6).Text
                CostCenter = Normalize-Text $Worksheet.Cells.Item($row, 7).Text
                Asiento = Normalize-Text $Worksheet.Cells.Item($row, 10).Text
            }) | Out-Null
        }
    } finally {
        if ($null -ne $usedRange) {
            [void][Runtime.Interopservices.Marshal]::ReleaseComObject($usedRange)
        }
    }

    return $prototypes.ToArray()
}

function Read-PrecontCostos2Prototypes {
    param([object]$Worksheet)

    $prototypes = New-Object System.Collections.Generic.List[object]
    $usedRange = $null
    try {
        $usedRange = $Worksheet.UsedRange
        $lastRow = [int]($usedRange.Row + $usedRange.Rows.Count - 1)
        for ($row = 2; $row -le $lastRow; $row++) {
            $accountText = Normalize-Text $Worksheet.Cells.Item($row, 4).Text
            $accountDigits = ($accountText -replace '[^0-9]', '')
            $account = if ($accountDigits.Length -ge 10) { $accountDigits.PadLeft(12, '0') } else { $accountText }
            if ($account -eq '') {
                continue
            }

            $prototypes.Add([pscustomobject]@{
                TemplateRow = $row
                Ag = Normalize-Text $Worksheet.Cells.Item($row, 2).Text
                Line = Normalize-Text $Worksheet.Cells.Item($row, 3).Text
                Account = $account
                Number = Normalize-Text $Worksheet.Cells.Item($row, 5).Text
                Description = Normalize-Text $Worksheet.Cells.Item($row, 6).Text
                CostCenter = Normalize-Text $Worksheet.Cells.Item($row, 7).Text
                Asiento = Normalize-Text $Worksheet.Cells.Item($row, 10).Text
            }) | Out-Null
        }
    } finally {
        if ($null -ne $usedRange) {
            [void][Runtime.Interopservices.Marshal]::ReleaseComObject($usedRange)
        }
    }

    return $prototypes.ToArray()
}

function New-PrecontVentasGeneratedRows {
    param(
        [object[]]$Prototypes,
        [object[]]$RowsPosting,
        [object[]]$MayorRows,
        [object[]]$PxRows
    )

    $generated = New-Object System.Collections.Generic.List[object]

    $invoiceTotalsByDocCenter = @{
        'FA' = @{}
        'FC' = @{}
    }
    $ivaTotalsByDocCenter = @{
        'FA' = @{}
        'FC' = @{}
    }
    $cashNotesByCenter = @{}
    $invoiceDocTypeByDocument = @{}
    $invoiceInfoByDocument = @{}
    $pxGrossByDocCenter = @{
        'FA' = @{}
        'FC' = @{}
    }
    $pxDiscountByDocCenter = @{
        'FA' = @{}
        'FC' = @{}
    }

    foreach ($row in @($RowsPosting | Where-Object { $_.DocType -in @('FA', 'FC') })) {
        $docType = (Normalize-Text $row.DocType).ToUpperInvariant()
        $center = Get-NormalizedCenterCode $row.Center
        if ($center -eq '') {
            $center = '00'
        }

        $amounts = Get-Invoice-SourceAmounts -Row $row
        if (-not $invoiceTotalsByDocCenter[$docType].ContainsKey($center)) {
            $invoiceTotalsByDocCenter[$docType][$center] = [double]0
        }
        if (-not $ivaTotalsByDocCenter[$docType].ContainsKey($center)) {
            $ivaTotalsByDocCenter[$docType][$center] = [double]0
        }

        $invoiceTotalsByDocCenter[$docType][$center] = [double]$invoiceTotalsByDocCenter[$docType][$center] + [double]$amounts.Total
        $ivaTotalsByDocCenter[$docType][$center] = [double]$ivaTotalsByDocCenter[$docType][$center] + [double]$amounts.Iva

        $documentKey = Trim-Document $row.DocumentTrim
        if ($documentKey -ne '' -and -not $invoiceDocTypeByDocument.ContainsKey($documentKey)) {
            $invoiceDocTypeByDocument[$documentKey] = $docType
        }
        if ($documentKey -ne '' -and -not $invoiceInfoByDocument.ContainsKey($documentKey)) {
            $invoiceInfoByDocument[$documentKey] = [pscustomobject]@{
                DocType = $docType
                Center = $center
            }
        }
    }

    foreach ($row in @($RowsPosting | Where-Object { $_.DocType -in @('DC', 'DE') })) {
        $center = Get-NormalizedCenterCode $row.Center
        if ($center -eq '') {
            $center = '00'
        }

        $affectedKey = Trim-Document $row.AffectedDocumentTrim
        if ($affectedKey -eq '' -or -not $invoiceDocTypeByDocument.ContainsKey($affectedKey)) {
            continue
        }

        if ($invoiceDocTypeByDocument[$affectedKey] -ne 'FA') {
            continue
        }

        $amounts = Get-Note-SourceAmounts -Row $row
        if (-not $cashNotesByCenter.ContainsKey($center)) {
            $cashNotesByCenter[$center] = [double]0
        }

        $cashNotesByCenter[$center] = [double]$cashNotesByCenter[$center] + [double]$amounts.Total
    }

    foreach ($pxRow in @(Convert-PxRowsToDetailRows -Rows $PxRows)) {
        $documentKey = Trim-Document $pxRow.Factura
        if ($documentKey -eq '' -or -not $invoiceInfoByDocument.ContainsKey($documentKey)) {
            continue
        }

        $invoiceInfo = $invoiceInfoByDocument[$documentKey]
        $docType = (Normalize-Text $invoiceInfo.DocType).ToUpperInvariant()
        if ($docType -notin @('FA', 'FC')) {
            continue
        }

        $center = Get-NormalizedCenterCode $invoiceInfo.Center
        if ($center -eq '') {
            $center = '00'
        }

        if (-not $pxGrossByDocCenter[$docType].ContainsKey($center)) {
            $pxGrossByDocCenter[$docType][$center] = [double]0
        }
        if (-not $pxDiscountByDocCenter[$docType].ContainsKey($center)) {
            $pxDiscountByDocCenter[$docType][$center] = [double]0
        }

        $pxGrossByDocCenter[$docType][$center] = [double]$pxGrossByDocCenter[$docType][$center] + [double](To-Number $pxRow.PvpBruto)
        $pxDiscountByDocCenter[$docType][$center] = [double]$pxDiscountByDocCenter[$docType][$center] + [double](To-Number $pxRow.DescValor)
    }

    $salesMayorByAccount = @{}
    foreach ($mayorRow in @($MayorRows)) {
        $account = Get-CompactAccountCode $mayorRow.account
        if ($account -eq '') {
            continue
        }

        if (-not $salesMayorByAccount.ContainsKey($account)) {
            $salesMayorByAccount[$account] = New-Object System.Collections.Generic.List[object]
        }

        $salesMayorByAccount[$account].Add($mayorRow) | Out-Null
    }

    $clientFaPrototypes = @($Prototypes | Where-Object { $_.Doc -eq 'FA' -and $_.Description -match '^CLIENTES SERV' })
    $clientFcPrototypes = @($Prototypes | Where-Object { $_.Doc -eq 'FC' -and $_.Description -match '^CLIENTES SERV' })
    $clientCaPrototypes = @($Prototypes | Where-Object { $_.Doc -eq 'CA' -and $_.Description -match '^CLIENTES SERV' })
    $ivaFaPrototypes = @($Prototypes | Where-Object { $_.Doc -eq 'FA' -and $_.Description -match 'IVA' })
    $ivaFcPrototypes = @($Prototypes | Where-Object { $_.Doc -eq 'FC' -and $_.Description -match 'IVA' })
    $reserveCaPrototypes = @($Prototypes | Where-Object { $_.Doc -eq 'CA' -and $_.Description -match 'RESERVA' })
    $liquidarFaCreditPrototypes = @($Prototypes | Where-Object { $_.Doc -eq 'FA' -and $_.Account -match '^020120\d{2}0002$' })
    $liquidarFaDiscountPrototypes = @($Prototypes | Where-Object { $_.Doc -eq 'FA' -and $_.Account -match '^020120\d{2}0004$' })
    $liquidarFcCreditPrototypes = @($Prototypes | Where-Object { $_.Doc -eq 'FC' -and $_.Account -match '^020120\d{2}0002$' })
    $liquidarFcDiscountPrototypes = @($Prototypes | Where-Object { $_.Doc -eq 'FC' -and $_.Account -match '^020120\d{2}0004$' })

    Add-TemplateAggregateRows -Target $generated -Prototypes $clientFaPrototypes -TotalsByCenter $invoiceTotalsByDocCenter['FA'] -AmountSide 'Debit'
    Add-TemplateAggregateRows -Target $generated -Prototypes $ivaFaPrototypes -TotalsByCenter $ivaTotalsByDocCenter['FA'] -AmountSide 'Credit'
    Add-TemplateAggregateRows -Target $generated -Prototypes $clientFcPrototypes -TotalsByCenter $invoiceTotalsByDocCenter['FC'] -AmountSide 'Debit'
    Add-TemplateAggregateRows -Target $generated -Prototypes $ivaFcPrototypes -TotalsByCenter $ivaTotalsByDocCenter['FC'] -AmountSide 'Credit'
    Add-TemplateAggregateRows -Target $generated -Prototypes $clientCaPrototypes -TotalsByCenter $cashNotesByCenter -AmountSide 'Credit'
    Add-TemplateAggregateRows -Target $generated -Prototypes $reserveCaPrototypes -TotalsByCenter $cashNotesByCenter -AmountSide 'Debit' -EnsureAtLeastOneRow
    Add-TemplateAggregateRows -Target $generated -Prototypes $liquidarFaCreditPrototypes -TotalsByCenter $pxGrossByDocCenter['FA'] -AmountSide 'Credit' -EnsureAtLeastOneRow
    Add-TemplateAggregateRows -Target $generated -Prototypes $liquidarFaDiscountPrototypes -TotalsByCenter $pxDiscountByDocCenter['FA'] -AmountSide 'Debit' -EnsureAtLeastOneRow
    Add-TemplateAggregateRows -Target $generated -Prototypes $liquidarFcCreditPrototypes -TotalsByCenter $pxGrossByDocCenter['FC'] -AmountSide 'Credit' -EnsureAtLeastOneRow
    Add-TemplateAggregateRows -Target $generated -Prototypes $liquidarFcDiscountPrototypes -TotalsByCenter $pxDiscountByDocCenter['FC'] -AmountSide 'Debit' -EnsureAtLeastOneRow

    $criticalAccounts = @(
        $Prototypes | Where-Object { $_.Account -match '^040101\d{2}0001$' } | Select-Object -ExpandProperty Account
        $Prototypes | Where-Object { $_.Account -match '^040101\d{2}0003$' } | Select-Object -ExpandProperty Account
        $Prototypes | Where-Object { $_.Account -match '^040101\d{2}0010$' } | Select-Object -ExpandProperty Account
        $Prototypes | Where-Object { $_.Account -match '^040101\d{2}0012$' } | Select-Object -ExpandProperty Account
        $Prototypes | Where-Object { $_.Account -match '^040101\d{2}0014$' } | Select-Object -ExpandProperty Account
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | ForEach-Object { [string]$_ } | Sort-Object -Unique

    foreach ($account in $criticalAccounts) {
        $accountPrototypes = @($Prototypes | Where-Object { $_.Account -eq $account } | Sort-Object @{ Expression = { [int]$_.TemplateRow } })
        $mayorGroup = if ($salesMayorByAccount.ContainsKey($account)) {
            @($salesMayorByAccount[$account].ToArray())
        } else {
            @()
        }
        Add-TemplateSequentialRowsFromMayor -Target $generated -Prototypes $accountPrototypes -MayorRows $mayorGroup -EnsureAtLeastOneRow
    }

    return $generated.ToArray()
}

function Write-PrecontVentasGeneratedRows {
    param(
        [object]$Worksheet,
        [object[]]$Rows
    )

    Clear-Worksheet-RangeContents -Worksheet $Worksheet -StartColumn 'B' -StartRow 2 -EndColumn 'J' -EndRow 1412
    if (@($Rows).Count -gt 0) {
        $sheetRows = New-Object System.Collections.Generic.List[object]
        foreach ($row in @($Rows)) {
            $sheetRows.Add(@(
                (Get-Excel-TextLiteral $row.Ag),
                (Get-Excel-TextLiteral $row.Doc),
                (Get-Excel-TextLiteral $row.Line),
                (Get-Excel-TextLiteral $row.Account),
                (Get-Excel-TextLiteral $row.Description),
                (Get-Excel-TextLiteral $row.CostCenter),
                (Get-NumericMatrixValue -Value $row.Debit -BlankIfZero),
                (Get-NumericMatrixValue -Value $row.Credit -BlankIfZero),
                (Get-Excel-TextLiteral $row.Asiento)
            )) | Out-Null
        }

        Write-Rows-ToWorksheet -Worksheet $Worksheet -Rows $sheetRows.ToArray() -StartRow 2 -StartColumn 2
    }

    return [int]@($Rows).Count
}

function Get-BrandCostMetrics {
    param([object[]]$Rows)

    $metrics = @{
        Costo = [double]0
        CostoLubricantes = [double]0
        CostoAccesorios = [double]0
        CostoRepuestos = [double]0
        CostoPintura = [double]0
        CostoSubconNc = [double]0
    }

    foreach ($row in @($Rows)) {
        $metrics.Costo = [double]$metrics.Costo + [double](To-Number $row.Costo)
        $metrics.CostoLubricantes = [double]$metrics.CostoLubricantes + [double](To-Number $row.CostoLubricantes)
        $metrics.CostoAccesorios = [double]$metrics.CostoAccesorios + [double](To-Number $row.CostoAccesorios)
        $metrics.CostoRepuestos = [double]$metrics.CostoRepuestos + [double](To-Number $row.CostoRepuestos)
        $metrics.CostoPintura = [double]$metrics.CostoPintura + [double](To-Number $row.CostoPintura)
        $metrics.CostoSubconNc = [double]$metrics.CostoSubconNc + [double](To-Number $row.CostoSubconNc)
    }

    foreach ($key in @($metrics.Keys)) {
        $metrics[$key] = [double](Round-Amount $metrics[$key])
    }

    return $metrics
}

function Write-PrecontCostos2GeneratedRows {
    param(
        [object]$Worksheet,
        [object[]]$Prototypes,
        [hashtable]$Metrics
    )

    $rowDefinitions = @(
        @{ Account = '050201010001'; Amount = [double]$Metrics.CostoRepuestos },
        @{ Account = '050201010002'; Amount = [double]$Metrics.CostoLubricantes },
        @{ Account = '050201010003'; Amount = [double]([double]$Metrics.Costo + [double]$Metrics.CostoSubconNc) },
        @{ Account = '050201010004'; Amount = [double]$Metrics.CostoPintura },
        @{ Account = '050201010005'; Amount = [double]0 },
        @{ Account = '050201010008'; Amount = [double]$Metrics.CostoAccesorios }
    )

    Clear-Worksheet-RangeContents -Worksheet $Worksheet -StartColumn 'B' -StartRow 2 -EndColumn 'J' -EndRow 51
    $sheetRows = New-Object System.Collections.Generic.List[object]
    foreach ($definition in $rowDefinitions) {
        $prototype = @($Prototypes | Where-Object { $_.Account -eq $definition.Account } | Select-Object -First 1)[0]
        if ($null -eq $prototype) {
            continue
        }

        $sheetRows.Add(@(
            (Get-Excel-TextLiteral $prototype.Ag),
            (Get-Excel-TextLiteral $prototype.Line),
            (Get-Excel-TextLiteral $prototype.Account),
            (Get-Excel-TextLiteral $prototype.Number),
            (Get-Excel-TextLiteral $prototype.Description),
            (Get-Excel-TextLiteral $prototype.CostCenter),
            (Get-NumericMatrixValue -Value $definition.Amount -BlankIfZero),
            $null,
            (Get-Excel-TextLiteral $prototype.Asiento)
        )) | Out-Null
    }

    if ($sheetRows.Count -gt 0) {
        Write-Rows-ToWorksheet -Worksheet $Worksheet -Rows $sheetRows.ToArray() -StartRow 2 -StartColumn 2
    }

    return [int]$sheetRows.Count
}

function Read-EstadisticasSeedRows {
    param(
        [object]$Worksheet,
        [int[]]$RowNumbers
    )

    $seedRows = @{}
    foreach ($rowIndex in @($RowNumbers)) {
        $seedRows[$rowIndex] = [pscustomobject]@{
            Account = Normalize-Text $Worksheet.Cells.Item($rowIndex, 1).Text
            Description = Normalize-Text $Worksheet.Cells.Item($rowIndex, 2).Text
            Mod = Normalize-Text $Worksheet.Cells.Item($rowIndex, 5).Text
            Asiento = Normalize-Text $Worksheet.Cells.Item($rowIndex, 6).Text
            Detalle = Normalize-Text $Worksheet.Cells.Item($rowIndex, 8).Text
        }
    }

    return $seedRows
}

function Write-EstadisticasGeneratedRows {
    param(
        [object]$Worksheet,
        [hashtable]$Metrics,
        [double]$PeriodDateValue
    )

    $rowNumbers = @(6, 131, 222, 290, 365)
    $seedRows = Read-EstadisticasSeedRows -Worksheet $Worksheet -RowNumbers $rowNumbers

    $clearRanges = @(
        @{ StartRow = 6; EndRow = 128; EndColumn = 'J' },
        @{ StartRow = 131; EndRow = 218; EndColumn = 'J' },
        @{ StartRow = 222; EndRow = 287; EndColumn = 'J' },
        @{ StartRow = 290; EndRow = 359; EndColumn = 'J' },
        @{ StartRow = 365; EndRow = 442; EndColumn = 'J' },
        @{ StartRow = 446; EndRow = 586; EndColumn = 'K' }
    )
    foreach ($range in $clearRanges) {
        $estadRangeRef = $Worksheet.Range("A$($range.StartRow):$($range.EndColumn)$($range.EndRow)")
        try {
            $estadConstants = $estadRangeRef.SpecialCells(2)
            $null = $estadConstants.ClearContents()
            [void][Runtime.Interopservices.Marshal]::ReleaseComObject($estadConstants)
        } catch {}
        [void][Runtime.Interopservices.Marshal]::ReleaseComObject($estadRangeRef)
    }

    $blockRows = @(
        @{ Row = 6; Amount = [double]$Metrics.CostoRepuestos },
        @{ Row = 131; Amount = [double]$Metrics.CostoLubricantes },
        @{ Row = 222; Amount = [double]([double]$Metrics.Costo + [double]$Metrics.CostoSubconNc) },
        @{ Row = 290; Amount = [double]$Metrics.CostoPintura },
        @{ Row = 365; Amount = [double]$Metrics.CostoAccesorios }
    )

    $sheetRows = New-Object System.Collections.Generic.List[object]
    foreach ($definition in $blockRows) {
        $rowIndex = [int]$definition.Row
        $seed = $seedRows[$rowIndex]
        $account = ''
        $description = ''
        $mod = ''
        $asiento = ''
        $detalle = ''
        if ($null -ne $seed) {
            $account = $seed.Account
            $description = $seed.Description
            $mod = $seed.Mod
            $asiento = $seed.Asiento
            $detalle = $seed.Detalle
        }

        $dateMode = Get-DateColumnWriteMode -Worksheet $Worksheet -Row $rowIndex -Column 4
        $sheetRows.Add([pscustomobject]@{
            RowIndex = $rowIndex
            Values = @(
                (Get-Excel-TextLiteral $account),
                (Get-Excel-TextLiteral $description),
                'N',
                (Get-DateMatrixValue -Value $PeriodDateValue -Mode $dateMode),
                (Get-Excel-TextLiteral $(if ($mod -eq '') { 'COSSE' } else { $mod })),
                (Get-Excel-TextLiteral $asiento),
                $null,
                (Get-Excel-TextLiteral $detalle),
                (Get-NumericMatrixValue -Value $definition.Amount -BlankIfZero),
                $null
            )
        }) | Out-Null
    }

    foreach ($sheetRow in $sheetRows) {
        Write-Rows-ToWorksheet -Worksheet $Worksheet -Rows @($sheetRow.Values) -StartRow ([int]$sheetRow.RowIndex) -StartColumn 1
    }
}

function Write-CostoGeneratedRows {
    param(
        [object]$Worksheet,
        [hashtable]$Metrics,
        [double]$PeriodDateValue
    )

    Clear-Worksheet-RangeContents -Worksheet $Worksheet -StartColumn 'A' -StartRow 6 -EndColumn 'J' -EndRow 141

    $rows = @(
        @{ Row = 6; Detail = 'REPUESTOS'; Amount = [double]$Metrics.CostoRepuestos },
        @{ Row = 7; Detail = 'INSUMOS Y LUBRICANTES'; Amount = [double]$Metrics.CostoLubricantes },
        @{ Row = 8; Detail = 'SUBCONTRATOS'; Amount = [double]([double]$Metrics.Costo + [double]$Metrics.CostoSubconNc) },
        @{ Row = 9; Detail = 'ACCESORIOS'; Amount = [double]$Metrics.CostoAccesorios }
    )

    $sheetRows = New-Object System.Collections.Generic.List[object]
    foreach ($definition in $rows) {
        $rowIndex = [int]$definition.Row
        $dateMode = Get-DateColumnWriteMode -Worksheet $Worksheet -Row $rowIndex -Column 3
        $sheetRows.Add([pscustomobject]@{
            RowIndex = $rowIndex
            Values = @(
                '05.01.01.01.0005',
                'COSTO DE VENTAS-SERVICIO',
                (Get-DateMatrixValue -Value $PeriodDateValue -Mode $dateMode),
                'COSSE',
                (Get-Excel-TextLiteral ([string]($rowIndex - 5))),
                $null,
                (Get-Excel-TextLiteral $definition.Detail),
                (Get-NumericMatrixValue -Value $definition.Amount -BlankIfZero),
                $null
            )
        }) | Out-Null
    }

    foreach ($sheetRow in $sheetRows) {
        Write-Rows-ToWorksheet -Worksheet $Worksheet -Rows @($sheetRow.Values) -StartRow ([int]$sheetRow.RowIndex) -StartColumn 1
    }
}

function Clear-PrecontCostosWorksheet {
    param([object]$Worksheet)

    Clear-Worksheet-RangeContents -Worksheet $Worksheet -StartColumn 'B' -StartRow 2 -EndColumn 'J' -EndRow 922
    Clear-Worksheet-RangeContents -Worksheet $Worksheet -StartColumn 'P' -StartRow 1 -EndColumn 'T' -EndRow 25
}

function Refresh-WorksheetPivotTablesSafe {
    param(
        [object]$Worksheet,
        [string]$Label
    )

    if (-not $script:PivotRefreshEnabled) {
        Write-Output ("INFO|pivot_refresh_skipped|{0}|mode=fast" -f $Label)
        return
    }

    try {
        $pivotTables = $Worksheet.PivotTables()
        $count = [int]$pivotTables.Count
        for ($index = 1; $index -le $count; $index++) {
            $pivotTable = $pivotTables.Item($index)
            try {
                $sourceData = Normalize-Text $pivotTable.SourceData
                if ($sourceData -match '^https?://' -or $sourceData -match 'sharepoint\.com') {
                    Write-Output ("WARN|pivot_refresh_skipped_external|{0}|{1}" -f $Label, $pivotTable.Name)
                    continue
                }

                $hadAutoFormat = $false
                try { $hadAutoFormat = [bool]$pivotTable.HasAutoFormat } catch {}
                if ($hadAutoFormat) { try { $pivotTable.HasAutoFormat = $false } catch {} }
                $null = $pivotTable.RefreshTable()
                if ($hadAutoFormat) { try { $pivotTable.HasAutoFormat = $true } catch {} }
            } finally {
                if ($null -ne $pivotTable) {
                    [void][Runtime.Interopservices.Marshal]::ReleaseComObject($pivotTable)
                }
            }
        }
        [void][Runtime.Interopservices.Marshal]::ReleaseComObject($pivotTables)
    } catch {
        throw "No se pudo refrescar la tabla dinamica de ${Label}: $($_.Exception.Message)"
    }
}

function Recalculate-WorksheetSafe {
    param(
        [object]$Worksheet,
        [string]$Label
    )

    try {
        $Worksheet.Calculate() | Out-Null
    } catch {
        Write-Output ("WARN|recalc_failed|{0}|{1}" -f $Label, $_.Exception.Message)
    }
}

function Get-Display-Cedula {
    param(
        [pscustomobject]$Lookup,
        [object]$Fallback
    )

    if ($null -ne $Lookup -and (Normalize-Text $Lookup.Cedula) -ne '') {
        return $Lookup.Cedula
    }

    return $Fallback
}

function Get-Display-Customer {
    param(
        [pscustomobject]$Lookup,
        [object]$Fallback
    )

    if ($null -ne $Lookup -and (Normalize-Text $Lookup.Customer) -ne '') {
        return $Lookup.Customer
    }

    return $Fallback
}

function Test-ObfuscatedText {
    param([object]$Value)

    $text = Normalize-Text $Value
    if ($text -eq '') {
        return $false
    }

    if ($text -match 'X{3,}') {
        return $true
    }

    if ($text -match '[~`!@#$%^&*{}[\]|<>]') {
        return $true
    }

    return $false
}

function Get-PreferredVisibleText {
    param(
        [object]$Primary,
        [object]$Secondary,
        [object]$Tertiary = ''
    )

    $primaryText = Normalize-Text $Primary
    if ($primaryText -ne '' -and -not (Test-ObfuscatedText $primaryText)) {
        return $Primary
    }

    $secondaryText = Normalize-Text $Secondary
    if ($secondaryText -ne '') {
        return $Secondary
    }

    $tertiaryText = Normalize-Text $Tertiary
    if ($tertiaryText -ne '') {
        return $Tertiary
    }

    return $Primary
}

function Get-PreferredSourceText {
    param(
        [object]$SourceRaw,
        [object]$SourceNormalized = '',
        [object]$LookupValue = ''
    )

    $sourceRawText = Normalize-Text $SourceRaw
    # La plantilla solo rellena faltantes reales; no debe pisar lo que si vino en el upload.
    if ($sourceRawText -ne '') {
        return $SourceRaw
    }

    $sourceNormalizedText = Normalize-Text $SourceNormalized
    if ($sourceNormalizedText -ne '') {
        return $SourceNormalized
    }

    if (-not $script:AllowTemplateDataFallback) {
        return $SourceNormalized
    }

    $lookupText = Normalize-Text $LookupValue
    if ($lookupText -ne '') {
        return $LookupValue
    }

    return $SourceNormalized
}

function Get-PreferredLookupText {
    param(
        [object]$LookupValue,
        [object]$SourceRaw = '',
        [object]$SourceNormalized = ''
    )

    return Get-PreferredSourceText `
        -SourceRaw $SourceRaw `
        -SourceNormalized $SourceNormalized `
        -LookupValue $LookupValue
}

function Get-PreferredSourceDate {
    param(
        [object]$SourceValue,
        [object]$LookupValue = $null
    )

    if ($null -ne $SourceValue -and $SourceValue -ne '') {
        return $SourceValue
    }

    if (-not $script:AllowTemplateDataFallback) {
        return $null
    }

    return $LookupValue
}

function Get-LookupDefaultText {
    param(
        [hashtable]$Lookups,
        [string]$Section,
        [string]$Field
    )

    if (-not $script:AllowTemplateDataFallback) {
        return ''
    }

    if ($null -eq $Lookups -or -not $Lookups.ContainsKey('Defaults')) {
        return ''
    }

    $defaults = $Lookups.Defaults
    if ($null -eq $defaults -or -not $defaults.ContainsKey($Section)) {
        return ''
    }

    $sectionDefaults = $defaults[$Section]
    if ($null -eq $sectionDefaults -or -not $sectionDefaults.ContainsKey($Field)) {
        return ''
    }

    return Normalize-Text $sectionDefaults[$Field]
}

function Resolve-TemplateGarExt {
    param(
        [object]$LookupValue,
        [object]$SourceRaw = '',
        [object]$SourceNormalized = '',
        [object]$TemplateDefault = ''
    )

    $sourceValue = Get-PreferredSourceText -SourceRaw $SourceRaw -SourceNormalized $SourceNormalized
    if (-not $script:AllowTemplateDataFallback) {
        return $sourceValue
    }

    $sourceText = (Normalize-Text $sourceValue).ToUpperInvariant()
    $defaultText = Normalize-Text $TemplateDefault

    if ($sourceText -ne '') {
        if ($sourceText -in @('N', 'NO', '0', 'FALSE')) {
            return $defaultText
        }

        return $sourceValue
    }

    $lookupText = Normalize-Text $LookupValue
    if ($lookupText -eq '') {
        return $defaultText
    }

    if ($lookupText.ToUpperInvariant() -in @('N', 'NO', '0', 'FALSE')) {
        return $defaultText
    }

    return $LookupValue
}

function Get-Excel-TextLiteral {
    param([object]$Value)

    if ($null -eq $Value) {
        return ''
    }

    $text = [string]$Value
    if ($text -eq '') {
        return ''
    }

    if ($text.StartsWith("'")) {
        return $text
    }

    $firstChar = $text.Substring(0, 1)
    if ($firstChar -in @('=', '+', '-', '@')) {
        return "'" + $text
    }

    if ($text -match '^\d+$') {
        return "'" + $text
    }

    return $text
}

function Get-Invoice-Asiento {
    param([pscustomobject]$Row)

    $paymentMethod = (Normalize-Text $Row.FormaPago).ToUpperInvariant()
    if ($paymentMethod -like '*CRED*') {
        return 'C'
    }

    if ($paymentMethod -like '*CONT*' -or $paymentMethod -like '*EFEC*') {
        return 'E'
    }

    $payment = (Normalize-Text $Row.DocType).ToUpperInvariant()
    if ($payment -eq 'FC') {
        return 'C'
    }

    return 'E'
}

function Set-NumericOrBlankCell {
    param(
        [object]$Worksheet,
        [int]$Row,
        [int]$Column,
        [double]$Value
    )

    if ((Round-Amount ([Math]::Abs($Value))) -eq 0) {
        $null = $Worksheet.Cells.Item($Row, $Column).ClearContents()
        return
    }

    $null = $Worksheet.Cells.Item($Row, $Column).Value2 = [double]$Value
}

function Set-TemplateNumericCell {
    param(
        [object]$Worksheet,
        [int]$Row,
        [int]$Column,
        [object]$TemplateText,
        [double]$Value
    )

    if ((Normalize-Text $TemplateText) -eq '') {
        $null = $Worksheet.Cells.Item($Row, $Column).ClearContents()
        return
    }

    $null = $Worksheet.Cells.Item($Row, $Column).Value2 = [double]$Value
}

function Set-NumericCellSafe {
    param(
        [object]$Worksheet,
        [int]$Row,
        [int]$Column,
        [double]$Value,
        [switch]$BlankIfZero
    )

    $rounded = [double](Round-Amount $Value)
    if ($BlankIfZero -and (Round-Amount ([Math]::Abs($rounded))) -eq 0) {
        $null = $Worksheet.Cells.Item($Row, $Column).ClearContents()
        return
    }

    try {
        $null = $Worksheet.Cells.Item($Row, $Column).Value2 = $rounded
    } catch {
        $null = $Worksheet.Cells.Item($Row, $Column).Value2 = [string]$rounded
    }
}

function Get-IvaBuckets {
    param(
        [pscustomobject]$Row,
        [double]$NetBase
    )

    $ivaTotal = Round-Amount ([Math]::Abs((To-Number $Row.Iva)))
    $iva12 = [double]0
    $iva15 = [double]0

    $hasIva12 = $Row.PSObject.Properties.Name -contains 'Iva12'
    $hasIva15 = $Row.PSObject.Properties.Name -contains 'Iva15'
    if ($hasIva12) {
        $iva12 = Round-Amount ([Math]::Abs((To-Number $Row.Iva12)))
    }
    if ($hasIva15) {
        $iva15 = Round-Amount ([Math]::Abs((To-Number $Row.Iva15)))
    }

    if (($iva12 + $iva15) -gt 0) {
        if ($ivaTotal -eq 0) {
            $ivaTotal = Round-Amount ($iva12 + $iva15)
        }
    } elseif ($NetBase -gt 0 -and $ivaTotal -gt 0) {
        $rate = [double]($ivaTotal / $NetBase)
        if ($rate -ge 0.14) {
            $iva15 = $ivaTotal
        } elseif ($rate -ge 0.105) {
            $iva12 = $ivaTotal
        } else {
            $iva12 = $ivaTotal
        }
    }

    if (($iva12 + $iva15) -eq 0 -and $ivaTotal -gt 0) {
        $iva12 = $ivaTotal
    }

    return [pscustomobject]@{
        Total = [double]$ivaTotal
        Iva12 = [double]$iva12
        Iva15 = [double]$iva15
    }
}

function Get-Invoice-SourceAmounts {
    param([pscustomobject]$Row)

    $subtotalRaw = [Math]::Abs($Row.TotalManoObra) + [Math]::Abs($Row.TotalSubcontratos) + [Math]::Abs($Row.TotalInsumos) + [Math]::Abs($Row.TotalAccesorios) + [Math]::Abs($Row.TotalRepuestos)
    $subtotal = Round-Amount $subtotalRaw
    $discount = Round-Amount ([Math]::Abs($Row.NoteCredit))
    $netoConIva = Round-Amount ($subtotal - $discount)
    $ivaBuckets = Get-IvaBuckets -Row $Row -NetBase $netoConIva
    $ivaAmount = [double]$ivaBuckets.Total
    $interestAmount = Round-Amount ([Math]::Abs($Row.Interes))
    $totalAmount = Round-Amount ($netoConIva + $ivaAmount + $interestAmount)
    $netoIva0Value = if ((Round-Amount ([Math]::Abs($ivaAmount))) -eq 0) { $netoConIva } else { 0.0 }
    $iva12Value = [double]$ivaBuckets.Iva12
    $iva15Value = [double]$ivaBuckets.Iva15

    return [pscustomobject]@{
        Total = [double]$totalAmount
        Iva = [double]$ivaAmount
        Iva12 = [double]$iva12Value
        Iva15 = [double]$iva15Value
        Interest = [double]$interestAmount
        NetoConIva = [double]$netoConIva
        Discount = [double]$discount
        Subtotal = [double]$subtotal
        NetoIva0 = [double]$netoIva0Value
    }
}

function Get-Note-SourceAmounts {
    param([pscustomobject]$Row)

    $subtotalRaw = [Math]::Abs($Row.TotalManoObra) + [Math]::Abs($Row.TotalSubcontratos) + [Math]::Abs($Row.TotalInsumos) + [Math]::Abs($Row.TotalAccesorios) + [Math]::Abs($Row.TotalRepuestos)
    $subtotal = Round-Amount $subtotalRaw
    $discount = Round-Amount ([Math]::Abs($Row.NoteCredit))
    $netoConIva = Round-Amount ($subtotal - $discount)
    $ivaBuckets = Get-IvaBuckets -Row $Row -NetBase $netoConIva
    $ivaAmount = [double]$ivaBuckets.Total
    $interestAmount = Round-Amount ([Math]::Abs($Row.Interes))
    $totalAmount = Round-Amount ($netoConIva + $ivaAmount + $interestAmount)
    $netoSinIva = if ((Round-Amount ([Math]::Abs($ivaAmount))) -eq 0) { $netoConIva } else { 0.0 }
    $iva12Value = [double]$ivaBuckets.Iva12
    $iva15Value = [double]$ivaBuckets.Iva15
    $anticipo = Round-Amount ([Math]::Abs((To-Number $Row.Anticipo)))
    $neto = Round-Amount ([Math]::Abs((To-Number $Row.Neto)))
    if ($neto -eq 0.0) {
        $neto = Round-Amount ($totalAmount - $anticipo)
    }

    return [pscustomobject]@{
        Total = [double]$totalAmount
        Iva = [double]$ivaAmount
        Iva12 = [double]$iva12Value
        Iva15 = [double]$iva15Value
        Interest = [double]$interestAmount
        NetoConIva = [double]$netoConIva
        Discount = [double]$discount
        Subtotal = [double]$subtotal
        NetoSinIva = [double]$netoSinIva
        Anticipo = [double]$anticipo
        Neto = [double]$neto
    }
}

function Fill-RepVtas {
    param(
        [object]$Worksheet,
        [object[]]$Rows,
        [hashtable]$Lookups
    )

    Clear-OutputSheet -Worksheet $Worksheet -StartRow 15 -LastColumn 'AL'
    $targetRow = 15
    $writtenRows = New-Object System.Collections.Generic.List[object]
    $sheetRows = New-Object System.Collections.Generic.List[object]
    $factDateMode = Get-DateColumnWriteMode -Worksheet $Worksheet -Row 15 -Column 9
    $noteDateMode = Get-DateColumnWriteMode -Worksheet $Worksheet -Row 15 -Column 10
    $sortedRows = @(
        $Rows |
            Sort-Object @{ Expression = {
                $lookupRow = Find-RepVtasEntry -Store $Lookups.RepVtas -DocKey $_.DocumentTrim -OrderKey $_.Order
                if ($null -ne $lookupRow) {
                    return [int]$lookupRow.RowOrder
                }

                return 100000 + [int]$_.RowIndex
            } }
    )

    foreach ($row in $sortedRows) {
        try {
            Assert-NotCancelled 'rep_vtas'
            $repLookup = Find-RepVtasEntry -Store $Lookups.RepVtas -DocKey $row.DocumentTrim -OrderKey $row.Order
            $garExtDefault = Get-LookupDefaultText -Lookups $Lookups -Section 'RepVtas' -Field 'GarExt'
            $agencyValue = Get-PreferredSourceText -SourceRaw $row.AgencyRaw -SourceNormalized $row.Agency -LookupValue $(if ($null -ne $repLookup) { $repLookup.Agency } else { '' })
            $centerValue = Get-PreferredSourceText -SourceRaw $row.CenterRaw -SourceNormalized $row.Center -LookupValue $(if ($null -ne $repLookup) { $repLookup.Center } else { '' })
            $orderValue = Get-PreferredSourceText -SourceRaw $row.OrderRaw -SourceNormalized $row.Order -LookupValue $(if ($null -ne $repLookup) { $repLookup.Order } else { '' })
            $advisorValue = Get-PreferredSourceText -SourceRaw $row.AdvisorRaw -SourceNormalized $row.Advisor -LookupValue $(if ($null -ne $repLookup) { $repLookup.Advisor } else { '' })
            $lineValue = Get-PreferredSourceText -SourceRaw $row.LineRaw -SourceNormalized $row.Line -LookupValue $(if ($null -ne $repLookup) { $repLookup.Line } else { '' })
            $cedulaValue = Get-PreferredSourceText -SourceRaw $row.CedulaRaw -SourceNormalized $row.Cedula -LookupValue $(if ($null -ne $repLookup) { $repLookup.Cedula } else { '' })
            $customerValue = Get-PreferredSourceText -SourceRaw $row.CustomerRaw -SourceNormalized $row.Customer -LookupValue $(if ($null -ne $repLookup) { $repLookup.Customer } else { '' })
            $documentRawValue = Get-PreferredSourceText -SourceRaw $row.DocumentRaw -SourceNormalized $row.DocumentTrim -LookupValue $(if ($null -ne $repLookup) { $repLookup.DocumentRaw } else { '' })
            $factDateValue = Get-PreferredSourceDate -SourceValue $row.DateFactValue -LookupValue $(if ($null -ne $repLookup) { $repLookup.DateFactValue } else { $null })
            $noteDateValue = Get-PreferredSourceDate -SourceValue $row.DateNoteValue -LookupValue $(if ($null -ne $repLookup) { $repLookup.DateNoteValue } else { $null })
            if (($row.DocType -in @('DC', 'DE')) -and ($null -eq $noteDateValue -or $noteDateValue -eq '') -and $null -ne $factDateValue) {
                $noteDateValue = $factDateValue
            }
            $noteCreditSource = $row.NoteCredit
            $totalManoObraSource = $row.TotalManoObra
            $totalSubcontratosSource = $row.TotalSubcontratos
            $totalInsumosSource = $row.TotalInsumos
            $totalServicioSource = $row.TotalServicio
            $totalAccesoriosSource = $row.TotalAccesorios
            $totalRepuestosSource = $row.TotalRepuestos
            $interesSource = $row.Interes
            $ivaSource = $row.Iva
            $totalSource = $row.Total
            $costoSource = $row.Costo
            $costoLubricantesSource = $row.CostoLubricantes
            $costoAccesoriosSource = $row.CostoAccesorios
            $costoRepuestosSource = $row.CostoRepuestos
            $costoPinturaSource = $row.CostoPintura
            $costoSubconNcSource = $row.CostoSubconNc

            $noteCreditValue = [double](Round-Amount $noteCreditSource)
            $totalManoObraValue = [double](Round-Amount $totalManoObraSource)
            $totalSubcontratosValue = [double](Round-Amount $totalSubcontratosSource)
            $totalInsumosValue = [double](Round-Amount $totalInsumosSource)
            $totalServicioValue = [double](Round-Amount $totalServicioSource)
            $totalAccesoriosValue = [double](Round-Amount $totalAccesoriosSource)
            $totalRepuestosValue = [double](Round-Amount $totalRepuestosSource)
            $interesValue = [double](Round-Amount $interesSource)
            $ivaValue = [double](Round-Amount $ivaSource)
            $totalValue = [double](Round-Amount $totalSource)
            $costoValue = [double](Round-Amount $costoSource)
            $costoLubricantesValue = [double](Round-Amount $costoLubricantesSource)
            $costoAccesoriosValue = [double](Round-Amount $costoAccesoriosSource)
            $costoRepuestosValue = [double](Round-Amount $costoRepuestosSource)
            $costoPinturaValue = [double](Round-Amount $costoPinturaSource)
            $costoSubconNcValue = [double](Round-Amount $costoSubconNcSource)
            $garExtValue = Resolve-TemplateGarExt -LookupValue $(if ($null -ne $repLookup) { $repLookup.GarExt } else { '' }) -SourceRaw $row.GarExtRaw -SourceNormalized $row.GarExt -TemplateDefault $garExtDefault
            $interestCellValue = if ((Round-Amount ([Math]::Abs($interesValue))) -eq 0) { $null } else { [double]$interesValue }
            $rowNumber = $targetRow

            $factDate = $factDateValue
            $noteDate = $noteDateValue
            $sheetRows.Add(@(
                (Get-Excel-TextLiteral $agencyValue),
                ("'" + $centerValue),
                (Get-Excel-TextLiteral $orderValue),
                (Get-Excel-TextLiteral $advisorValue),
                (Get-Excel-TextLiteral $lineValue),
                ("'" + $cedulaValue),
                (Get-Excel-TextLiteral $customerValue),
                ("'" + $documentRawValue),
                (Get-DateMatrixValue -Value $factDate -Mode $factDateMode),
                (Get-DateMatrixValue -Value $noteDate -Mode $noteDateMode),
                (Get-NumericMatrixValue -Value $noteCreditValue),
                (Get-NumericMatrixValue -Value $totalManoObraValue),
                (Get-NumericMatrixValue -Value $totalSubcontratosValue),
                (Get-NumericMatrixValue -Value $totalInsumosValue),
                (Get-NumericMatrixValue -Value $totalServicioValue),
                (Get-NumericMatrixValue -Value $totalAccesoriosValue),
                (Get-NumericMatrixValue -Value $totalRepuestosValue),
                (Get-NumericMatrixValue -Value $interesValue -BlankIfZero),
                (Get-NumericMatrixValue -Value $ivaValue),
                (Get-NumericMatrixValue -Value $totalValue),
                (Get-NumericMatrixValue -Value $costoValue),
                (Get-NumericMatrixValue -Value $costoLubricantesValue),
                (Get-NumericMatrixValue -Value $costoAccesoriosValue),
                (Get-NumericMatrixValue -Value $costoRepuestosValue),
                (Get-NumericMatrixValue -Value $costoPinturaValue),
                (Get-NumericMatrixValue -Value $costoSubconNcValue),
                (Get-Excel-TextLiteral $garExtValue)
            )) | Out-Null

            $writtenRows.Add([pscustomobject]@{
                RowNumber = $rowNumber
                Values = @{
                    1 = $agencyValue
                    2 = $centerValue
                    3 = $orderValue
                    4 = $advisorValue
                    5 = $lineValue
                    6 = $cedulaValue
                    7 = $customerValue
                    8 = $documentRawValue
                    9 = $factDate
                    10 = $noteDate
                    11 = [double]$noteCreditValue
                    12 = [double]$totalManoObraValue
                    13 = [double]$totalSubcontratosValue
                    14 = [double]$totalInsumosValue
                    15 = [double]$totalServicioValue
                    16 = [double]$totalAccesoriosValue
                    17 = [double]$totalRepuestosValue
                    18 = $interestCellValue
                    19 = [double]$ivaValue
                    20 = [double]$totalValue
                    21 = [double]$costoValue
                    22 = [double]$costoLubricantesValue
                    23 = [double]$costoAccesoriosValue
                    24 = [double]$costoRepuestosValue
                    25 = [double]$costoPinturaValue
                    26 = [double]$costoSubconNcValue
                    27 = $garExtValue
                }
            }) | Out-Null
            $targetRow++
        } catch {
            $failedLine = ''
            if ($null -ne $_.InvocationInfo -and $null -ne $_.InvocationInfo.Line) {
                $failedLine = $_.InvocationInfo.Line.Trim()
            }

            throw "REP VTAS doc=$($row.DocumentRaw) orden=$($row.Order) filaFuente=$($row.RowIndex): $($_.Exception.Message) :: $failedLine"
        }
    }

    if ($sheetRows.Count -gt 0) {
        Write-Rows-ToWorksheet -Worksheet $Worksheet -Rows $sheetRows.ToArray() -StartRow 15 -StartColumn 1
    }

    return [pscustomobject]@{
        RowCount = $writtenRows.Count
        Rows = $writtenRows.ToArray()
    }
}

function Fill-Invoices {
    param(
        [object]$Worksheet,
        [object[]]$Rows,
        [hashtable]$Lookups
    )

    Clear-OutputSheet -Worksheet $Worksheet -StartRow 17 -LastColumn 'S'
    $targetRow = 17
    $fallbackCount = 0
    $writtenRows = New-Object System.Collections.Generic.List[object]
    $sheetRows = New-Object System.Collections.Generic.List[object]
    $invoiceDateMode = Get-DateColumnWriteMode -Worksheet $Worksheet -Row 17 -Column 4
    $sortedRows = @(
        $Rows |
            Where-Object { $_.DocType -in @('FA', 'FC') } |
            Sort-Object `
                @{ Expression = { Normalize-Text $_.Agency } }, `
                @{ Expression = { Get-Document-SortValue $_.DocumentTrim } }, `
                @{ Expression = { Normalize-Text $_.Order } }
    )

    foreach ($row in $sortedRows) {
        try {
            Assert-NotCancelled 'facturas'
            $lookup = Find-LookupEntry -Store $Lookups.Invoice -DocKey $row.DocumentTrim -OrderKey $row.Order
            $garExtDefault = Get-LookupDefaultText -Lookups $Lookups -Section 'Invoice' -Field 'GarExt'
            if ($null -eq $lookup) {
                $fallbackCount++
            }

            $sourceAmounts = Get-Invoice-SourceAmounts -Row $row
            $totalAmount = [double]$sourceAmounts.Total
            $ivaAmount = [double]$sourceAmounts.Iva
            $iva12Amount = [double]$sourceAmounts.Iva12
            $iva15Amount = [double]$sourceAmounts.Iva15
            $interestAmount = [double]$sourceAmounts.Interest
            $netoConIva = [double]$sourceAmounts.NetoConIva
            $discount = [double]$sourceAmounts.Discount
            $subtotal = [double]$sourceAmounts.Subtotal
            $agencyValue = Get-PreferredLookupText -LookupValue $(if ($null -ne $lookup) { $lookup.Agency } else { '' }) -SourceRaw $row.AgencyRaw -SourceNormalized $row.Agency
            $seriesValue = Get-PreferredLookupText -LookupValue $(if ($null -ne $lookup) { $lookup.Series } else { '' }) -SourceRaw $row.SeriesRaw -SourceNormalized $row.Series
            $orderValue = Get-PreferredLookupText -LookupValue $(if ($null -ne $lookup) { $lookup.Order } else { '' }) -SourceRaw $row.OrderRaw -SourceNormalized $row.Order
            $cedulaValue = Get-PreferredLookupText -LookupValue $(if ($null -ne $lookup) { $lookup.Cedula } else { '' }) -SourceRaw $row.CedulaRaw -SourceNormalized $row.Cedula
            $customerValue = Get-PreferredLookupText -LookupValue $(if ($null -ne $lookup) { $lookup.Customer } else { '' }) -SourceRaw $row.CustomerRaw -SourceNormalized $row.Customer
            $netoIva0Value = [double]$sourceAmounts.NetoIva0
            $asientoValue = Get-PreferredSourceText -SourceRaw (Get-Invoice-Asiento -Row $row) -LookupValue $(if ($null -ne $lookup) { $lookup.Asiento } else { '' })
            $garExtValue = Resolve-TemplateGarExt -LookupValue $(if ($null -ne $lookup) { $lookup.GarExt } else { '' }) -SourceRaw $row.GarExtRaw -SourceNormalized $row.GarExt -TemplateDefault $garExtDefault
            $tvValue = Get-PreferredLookupText -LookupValue $(if ($null -ne $lookup) { $lookup.Tv } else { '' }) -SourceRaw $row.FormaPago -SourceNormalized $row.FormaPago
            $markerValue = 'N'

            $rowNumber = $targetRow

            $invoiceDate = Get-Date-Write-Value $row.DateFactValue
            $sheetRows.Add(@(
                (Get-Excel-TextLiteral $agencyValue),
                (Get-Excel-TextLiteral $seriesValue),
                $row.DocumentTrim,
                (Get-DateMatrixValue -Value $invoiceDate -Mode $invoiceDateMode),
                (Get-Excel-TextLiteral $orderValue),
                ("'" + $cedulaValue),
                (Get-Excel-TextLiteral $customerValue),
                (Get-NumericMatrixValue -Value $subtotal),
                (Get-NumericMatrixValue -Value $discount),
                (Get-NumericMatrixValue -Value $netoConIva),
                (Get-NumericMatrixValue -Value $netoIva0Value),
                (Get-NumericMatrixValue -Value $iva12Amount),
                (Get-NumericMatrixValue -Value $iva15Amount),
                (Get-NumericMatrixValue -Value $interestAmount),
                (Get-NumericMatrixValue -Value $totalAmount),
                $asientoValue,
                (Get-Excel-TextLiteral $garExtValue),
                (Get-Excel-TextLiteral $tvValue),
                (Get-Excel-TextLiteral $markerValue)
            )) | Out-Null

            $writtenRows.Add([pscustomobject]@{
                RowNumber = $rowNumber
                Values = @{
                    1 = $agencyValue
                    2 = $seriesValue
                    3 = $row.DocumentTrim
                    4 = $invoiceDate
                    5 = $orderValue
                    6 = $cedulaValue
                    7 = $customerValue
                    8 = [double]$subtotal
                    9 = [double]$discount
                    10 = [double]$netoConIva
                    11 = [double]$netoIva0Value
                    12 = [double]$iva12Amount
                    13 = [double]$iva15Amount
                    14 = [double]$interestAmount
                    15 = [double]$totalAmount
                    16 = $asientoValue
                    17 = $garExtValue
                    18 = $tvValue
                    19 = $markerValue
                }
            }) | Out-Null
            $targetRow++
        } catch {
            $failedLine = ''
            if ($null -ne $_.InvocationInfo -and $null -ne $_.InvocationInfo.Line) {
                $failedLine = $_.InvocationInfo.Line.Trim()
            }

            throw "FACTURA doc=$($row.DocumentTrim) orden=$($row.Order) filaFuente=$($row.RowIndex): $($_.Exception.Message) :: $failedLine"
        }
    }

    if ($sheetRows.Count -gt 0) {
        Write-Rows-ToWorksheet -Worksheet $Worksheet -Rows $sheetRows.ToArray() -StartRow 17 -StartColumn 1
    }

    return [pscustomobject]@{
        FallbackCount = $fallbackCount
        RowCount = $writtenRows.Count
        Rows = $writtenRows.ToArray()
    }
}

function Fill-Notes {
    param(
        [object]$Worksheet,
        [object[]]$Rows,
        [hashtable]$Lookups
    )

    Clear-OutputSheet -Worksheet $Worksheet -StartRow 11 -LastColumn 'U'
    $targetRow = 11
    $fallbackCount = 0
    $writtenRows = New-Object System.Collections.Generic.List[object]
    $sheetRows = New-Object System.Collections.Generic.List[object]
    $creditNoteDateMode = Get-DateColumnWriteMode -Worksheet $Worksheet -Row 11 -Column 3
    $sortedRows = @(
        $Rows |
            Where-Object { $_.DocType -in @('DC', 'DE') } |
            Sort-Object `
                @{ Expression = { Normalize-Text $_.Agency } }, `
                @{ Expression = { Get-Document-SortValue $_.DocumentTrim } }, `
                @{ Expression = { Normalize-Text $_.Order } }
    )

    foreach ($row in $sortedRows) {
        try {
            Assert-NotCancelled 'notas'
            $orderKey = Strip-Order-Suffix $row.Order
            $lookup = Find-LookupEntry -Store $Lookups.Note -DocKey $row.DocumentTrim -OrderKey $orderKey
            $garExtDefault = Get-LookupDefaultText -Lookups $Lookups -Section 'Note' -Field 'GarExt'
            if ($null -eq $lookup) {
                $fallbackCount++
            }

            $sourceAmounts = Get-Note-SourceAmounts -Row $row
            $totalAmount = [double]$sourceAmounts.Total
            $ivaAmount = [double]$sourceAmounts.Iva
            $iva12Amount = [double]$sourceAmounts.Iva12
            $iva15Amount = [double]$sourceAmounts.Iva15
            $interestAmount = [double]$sourceAmounts.Interest
            $netoConIva = [double]$sourceAmounts.NetoConIva
            $netoSinIva = [double]$sourceAmounts.NetoSinIva
            $discount = [double]$sourceAmounts.Discount
            $subtotal = [double]$sourceAmounts.Subtotal
            $anticipo = [double]$sourceAmounts.Anticipo
            $neto = [double]$sourceAmounts.Neto
            $agencyValue = Get-PreferredLookupText -LookupValue $(if ($null -ne $lookup) { $lookup.Agency } else { '' }) -SourceRaw $row.AgencyRaw -SourceNormalized $row.Agency
            $kindValue = if ($row.DocType -eq 'DE') {
                'CON'
            } else {
                Get-PreferredSourceText -SourceRaw 'CRE' -LookupValue $(if ($null -ne $lookup) { $lookup.Kind } else { '' })
            }
            $seriesValue = Get-PreferredLookupText -LookupValue $(if ($null -ne $lookup) { $lookup.Series } else { '' }) -SourceRaw $row.SeriesRaw -SourceNormalized $row.Series
            $invoiceValue = Get-PreferredSourceText -SourceRaw (Trim-Document $row.AffectedDocumentRaw) -SourceNormalized (Trim-Document $row.AffectedDocumentTrim) -LookupValue $(if ($null -ne $lookup) { $lookup.Invoice } else { '' })
            $orderValue = Get-PreferredSourceText -SourceRaw ([string]$row.OrderRaw) -SourceNormalized $(if ($orderKey -ne '') { $orderKey } else { Normalize-Text $row.Order }) -LookupValue $(if ($null -ne $lookup) { $lookup.Order } else { '' })
            $cedulaValue = Get-PreferredLookupText -LookupValue $(if ($null -ne $lookup) { $lookup.Cedula } else { '' }) -SourceRaw $row.CedulaRaw -SourceNormalized $row.Cedula
            $customerValue = Get-PreferredLookupText -LookupValue $(if ($null -ne $lookup) { $lookup.Customer } else { '' }) -SourceRaw $row.CustomerRaw -SourceNormalized $row.Customer
            $asientoValue = ''
            $garExtValue = Resolve-TemplateGarExt -LookupValue $(if ($null -ne $lookup) { $lookup.GarExt } else { '' }) -SourceRaw $row.GarExtRaw -SourceNormalized $row.GarExt -TemplateDefault $garExtDefault

            $rowNumber = $targetRow

            $creditNoteDate = Get-Date-Write-Value $row.DateNoteValue
            $sheetRows.Add(@(
                (Get-Excel-TextLiteral $agencyValue),
                $row.DocumentTrim,
                (Get-DateMatrixValue -Value $creditNoteDate -Mode $creditNoteDateMode),
                (Get-Excel-TextLiteral $kindValue),
                (Get-Excel-TextLiteral $seriesValue),
                (Get-Excel-TextLiteral $invoiceValue),
                (Get-Excel-TextLiteral $orderValue),
                ("'" + $cedulaValue),
                (Get-Excel-TextLiteral $customerValue),
                (Get-NumericMatrixValue -Value $subtotal),
                (Get-NumericMatrixValue -Value $discount),
                (Get-NumericMatrixValue -Value $netoSinIva),
                (Get-NumericMatrixValue -Value $netoConIva),
                (Get-NumericMatrixValue -Value $iva15Amount),
                (Get-NumericMatrixValue -Value $iva12Amount),
                (Get-NumericMatrixValue -Value $interestAmount),
                (Get-NumericMatrixValue -Value $totalAmount),
                (Get-NumericMatrixValue -Value $anticipo),
                (Get-NumericMatrixValue -Value $neto),
                (Get-Excel-TextLiteral $asientoValue),
                (Get-Excel-TextLiteral $garExtValue)
            )) | Out-Null

            $writtenRows.Add([pscustomobject]@{
                RowNumber = $rowNumber
                Values = @{
                    1 = $agencyValue
                    2 = $row.DocumentTrim
                    3 = $creditNoteDate
                    4 = $kindValue
                    5 = $seriesValue
                    6 = $invoiceValue
                    7 = $orderValue
                    8 = $cedulaValue
                    9 = $customerValue
                    10 = [double]$subtotal
                    11 = [double]$discount
                    12 = [double]$netoSinIva
                    13 = [double]$netoConIva
                    14 = [double]$iva15Amount
                    15 = [double]$iva12Amount
                    16 = [double]$interestAmount
                    17 = [double]$totalAmount
                    18 = [double]$anticipo
                    19 = [double]$neto
                    20 = $asientoValue
                    21 = $garExtValue
                }
            }) | Out-Null
            $targetRow++
        } catch {
            $failedLine = ''
            if ($null -ne $_.InvocationInfo -and $null -ne $_.InvocationInfo.Line) {
                $failedLine = $_.InvocationInfo.Line.Trim()
            }

            throw "NOTA doc=$($row.DocumentTrim) orden=$($row.Order) filaFuente=$($row.RowIndex): $($_.Exception.Message) :: $failedLine"
        }
    }

    if ($sheetRows.Count -gt 0) {
        Write-Rows-ToWorksheet -Worksheet $Worksheet -Rows $sheetRows.ToArray() -StartRow 11 -StartColumn 1
    }

    return [pscustomobject]@{
        FallbackCount = $fallbackCount
        RowCount = $writtenRows.Count
        Rows = $writtenRows.ToArray()
    }
}

function Validate-Services-BrandOutput {
    param(
        [object]$OutputWorkbook,
        [object]$TemplateWorkbook,
        [object]$RepVtasResult,
        [object]$InvoiceResult,
        [object]$NoteResult,
        [object]$MayorResult
    )

    if ($null -ne $MayorResult -and $null -ne $MayorResult.UnmappedAccounts -and @($MayorResult.UnmappedAccounts).Count -gt 0) {
        $unmappedSplit = Split-MayorUnmappedAccounts -Accounts @($MayorResult.UnmappedAccounts)
        if (@($unmappedSplit.Warning).Count -gt 0) {
            Write-Output ("WARN|mayor_unmapped_compatible|{0}|accounts={1}" -f $BrandKey, ((@($unmappedSplit.Warning) | Select-Object -Unique) -join ', '))
        }

        if (@($unmappedSplit.Fatal).Count -gt 0) {
            $accountList = (@($unmappedSplit.Fatal) | Where-Object { $_ -and $_ -ne '' } | Select-Object -Unique) -join ', '
            throw "El mayor contiene cuentas sin seccion equivalente en la plantilla: $accountList"
        }
    }

    $outputRepSheet = $null
    $outputNoteSheet = $null
    $outputRepVtasSheet = $null
    $outputPrecontVentasSheet = $null
    $outputPrecontCostosSheet = $null
    $templateRepSheet = $null
    $templateNoteSheet = $null
    $templateRepVtasSheet = $null
    try {
        $outputRepSheet = Get-Worksheet-Safe -Workbook $OutputWorkbook -CandidateNames @('REP FACTURACION', 'REP FACTURACIÃ“N')
        $outputNoteSheet = Get-Worksheet-Safe -Workbook $OutputWorkbook -CandidateNames @('NOTA DE CREDITO')
        $outputRepVtasSheet = Get-Worksheet-Safe -Workbook $OutputWorkbook -CandidateNames @('REP VTAS')
        $outputPrecontVentasSheet = Get-Worksheet-Safe -Workbook $OutputWorkbook -CandidateNames @('PrecontabilizacionVentas')
        $outputPrecontCostosSheet = Get-Worksheet-Safe -Workbook $OutputWorkbook -CandidateNames @('PrecontabilizacionCostos')

        $templateRepSheet = Get-Worksheet-Safe -Workbook $TemplateWorkbook -CandidateNames @('REP FACTURACION', 'REP FACTURACIÃ“N')
        $templateNoteSheet = Get-Worksheet-Safe -Workbook $TemplateWorkbook -CandidateNames @('NOTA DE CREDITO')
        $templateRepVtasSheet = Get-Worksheet-Safe -Workbook $TemplateWorkbook -CandidateNames @('REP VTAS')

        Validate-Services-WrittenRows `
            -OutputWorksheet $outputRepVtasSheet `
            -TemplateWorksheet $templateRepVtasSheet `
            -Rows $RepVtasResult.Rows `
            -Label 'REP VTAS' `
            -TextColumns @(1, 2, 3, 4, 5, 6, 7, 8, 27) `
            -NumericColumns @(11, 15, 19, 20, 21, 24, 26) `
            -DateColumns @(9, 10) `
            -FormatColumns @(9, 10) `
            -DocumentColumns @(8) `
            -TrailingBlankColumn 8 `
            -LastColumn 27

        Validate-Services-WrittenRows `
            -OutputWorksheet $outputRepSheet `
            -TemplateWorksheet $templateRepSheet `
            -Rows $InvoiceResult.Rows `
            -Label 'REP FACTURACION' `
            -TextColumns @(1, 2, 3, 5, 6, 7, 16, 17, 18, 19) `
            -NumericColumns @(8, 9, 10, 13, 14, 15) `
            -DateColumns @(4) `
            -FormatColumns @(4) `
            -DocumentColumns @(3) `
            -TrailingBlankColumn 3 `
            -LastColumn 19

        Validate-Services-WrittenRows `
            -OutputWorksheet $outputNoteSheet `
            -TemplateWorksheet $templateNoteSheet `
            -Rows $NoteResult.Rows `
            -Label 'NOTA DE CREDITO' `
            -TextColumns @(1, 2, 4, 5, 6, 7, 8, 9, 20, 21) `
            -NumericColumns @(10, 11, 13, 14, 17, 18, 19) `
            -DateColumns @(3) `
            -FormatColumns @(3) `
            -DocumentColumns @(2, 6) `
            -TrailingBlankColumn 2 `
            -LastColumn 21

        Assert-NoExcelErrorsInRange -Worksheet $outputRepSheet -StartRow 1 -EndRow 15 -LastColumn 21 -Label 'REP FACTURACION_CONTROL'
        Assert-NoExcelErrorsInRange -Worksheet $outputNoteSheet -StartRow 1 -EndRow 10 -LastColumn 21 -Label 'NOTA DE CREDITO_CONTROL'
        Assert-NoExcelErrorsInRange -Worksheet $outputRepVtasSheet -StartRow 1 -EndRow 8 -LastColumn 38 -Label 'REP VTAS_CONTROL'
        Assert-NoExcelErrorsInRange -Worksheet $outputPrecontVentasSheet -StartRow 1 -EndRow 12 -LastColumn 22 -Label 'PRECONTABILIZACION_VENTAS_CONTROL'
        # Esta hoja legacy se neutraliza a proposito porque la plantilla arrastra
        # movimientos historicos no alineados con la fuente mensual cargada.
        Assert-WorksheetAreaBlank -Worksheet $outputPrecontCostosSheet -StartRow 2 -EndRow 922 -StartColumn 2 -EndColumn 10 -Label 'PRECONTABILIZACION_COSTOS_LEGACY'
        Assert-WorksheetAreaBlank -Worksheet $outputPrecontCostosSheet -StartRow 1 -EndRow 25 -StartColumn 16 -EndColumn 20 -Label 'PRECONTABILIZACION_COSTOS_PIVOT'
    }
    finally {
        if ($null -ne $templateRepVtasSheet) {
            [void][Runtime.Interopservices.Marshal]::ReleaseComObject($templateRepVtasSheet)
        }
        if ($null -ne $templateNoteSheet) {
            [void][Runtime.Interopservices.Marshal]::ReleaseComObject($templateNoteSheet)
        }
        if ($null -ne $templateRepSheet) {
            [void][Runtime.Interopservices.Marshal]::ReleaseComObject($templateRepSheet)
        }
        if ($null -ne $outputRepVtasSheet) {
            [void][Runtime.Interopservices.Marshal]::ReleaseComObject($outputRepVtasSheet)
        }
        if ($null -ne $outputPrecontCostosSheet) {
            [void][Runtime.Interopservices.Marshal]::ReleaseComObject($outputPrecontCostosSheet)
        }
        if ($null -ne $outputPrecontVentasSheet) {
            [void][Runtime.Interopservices.Marshal]::ReleaseComObject($outputPrecontVentasSheet)
        }
        if ($null -ne $outputNoteSheet) {
            [void][Runtime.Interopservices.Marshal]::ReleaseComObject($outputNoteSheet)
        }
        if ($null -ne $outputRepSheet) {
            [void][Runtime.Interopservices.Marshal]::ReleaseComObject($outputRepSheet)
        }
    }
}

$primaryInput = if (-not [string]::IsNullOrWhiteSpace($RepVtasPath)) { $RepVtasPath } else { $InputPath }
$resolvedInputPath = Resolve-RequiredPath -Path $primaryInput -Label 'InputPath'
if (-not (Test-Path -LiteralPath $OutputDir)) {
    $null = New-Item -ItemType Directory -Path $OutputDir -Force
}
$resolvedOutputDir = (Resolve-Path -LiteralPath $OutputDir).Path
$resolvedTemplateDir = Resolve-RequiredPath -Path $TemplateDir -Label 'TemplateDir'

$stagingRoot = Join-Path $resolvedOutputDir '__staging_servicios_marcas'
if (-not (Test-Path -LiteralPath $stagingRoot)) {
    $null = New-Item -ItemType Directory -Path $stagingRoot -Force
}
$stagingDirectory = Join-Path $stagingRoot ([Guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $stagingDirectory -Force
$stagedInputPath = Join-Path $stagingDirectory ("input_source{0}" -f ([System.IO.Path]::GetExtension($resolvedInputPath)))
Copy-File-WithRetry -SourcePath $resolvedInputPath -DestinationPath $stagedInputPath
try {
$displaySourceRows = @()
$sourceRows = @()
$brandInputPaths = @{
    changan = @{
        FacturaPath = $FacturaChanganPath
        NotaPath = $NotaChanganPath
    }
    peug = @{
        FacturaPath = $FacturaPeugPath
        NotaPath = $NotaPeugPath
    }
    szk = @{
        FacturaPath = $FacturaSzkPath
        NotaPath = $NotaSzkPath
    }
    tyt = @{
        FacturaPath = $FacturaTytPath
        NotaPath = $NotaTytPath
    }
}
$displayRequested = $false
foreach ($paths in $brandInputPaths.Values) {
    if ((Normalize-Text $paths.FacturaPath) -ne '' -or (Normalize-Text $paths.NotaPath) -ne '') {
        $displayRequested = $true
        break
    }
}

if ($displayRequested) {
    $displaySourceRows = Build-SourceRows-FromBrandInputs -BrandFileMap $brandInputPaths -BrandKey $BrandKey
    if ($displaySourceRows.Count -gt 0) {
        Write-Output ("INFO|custom_source|rows={0}" -f $displaySourceRows.Count)
    } else {
        Write-Output 'WARN|custom_source_empty|falling_back_to_excel'
    }
}

if ($sourceRows.Count -eq 0) {
    $sourceRows = Read-SourceRows-FromFile -InputPath $stagedInputPath -WorkingDirectory $stagingDirectory
    if ($sourceRows.Count -gt 0 -and $displaySourceRows.Count -gt 0) {
        Write-Output ("INFO|source_priority|excel_primary|custom_support={0}" -f $displaySourceRows.Count)
    }
}

if ($sourceRows.Count -eq 0 -and $displaySourceRows.Count -gt 0) {
    $sourceRows = $displaySourceRows
}

if ($sourceRows.Count -eq 0) {
    throw 'Los archivos de entrada no contienen filas validas para procesar (custom y fallback vacios).'
}

$displayRows = if ($displaySourceRows.Count -gt 0) {
    Normalize-SourceRows -Rows $displaySourceRows -ConsolidateInvoiceDocuments
} else {
    Normalize-SourceRows -Rows $sourceRows -ConsolidateInvoiceDocuments
}
$repVtasRows = Normalize-SourceRows -Rows $sourceRows
$postingRows = Normalize-SourceRows -Rows $sourceRows -ConsolidateInvoiceDocuments
    if ($repVtasRows.Count -eq 0 -and $postingRows.Count -eq 0 -and $displayRows.Count -eq 0) {
        throw 'El archivo fuente no contiene filas validas para generar plantillas.'
    }
} catch {
    if (Test-Path -LiteralPath $stagingDirectory) {
        Remove-Item -LiteralPath $stagingDirectory -Recurse -Force -ErrorAction SilentlyContinue
    }
    throw
}

$mayorPaths = @{
    changan = $MayorChanganPath
    peug = $MayorPeugPath
    szk = $MayorSzkPath
    tyt = $MayorTytPath
}

$templateConfigs = @{
    changan = @{
        Label = 'CHANGAN'
        TemplatePath = (Join-Path $resolvedTemplateDir '11. Concili. Servicios CHANGAN  2026.xls')
    }
    peug = @{
        Label = 'PEUGEOT'
        TemplatePath = (Join-Path $resolvedTemplateDir '11. Concili. Servicios PEUG  2026.xls')
    }
    szk = @{
        Label = 'SUZUKI'
        TemplatePath = (Join-Path $resolvedTemplateDir '11. Concili. Servicios SZK  2026.xls')
    }
    tyt = @{
        Label = 'MATRIZ'
        TemplatePath = (Join-Path $resolvedTemplateDir '11. Concili. Servicios TYT 2026.xls')
    }
}

foreach ($config in $templateConfigs.Values) {
    if (-not (Test-Path -LiteralPath $config.TemplatePath)) {
        throw "No existe la plantilla base: $($config.TemplatePath)"
    }
}

$excelLock = Acquire-Excel-Automation-Lock -TimeoutSeconds 300
Stop-OrphanExcelProcesses -TimeoutSeconds 5
$_conflictPaths = @(
    $resolvedTemplateDir,
    $resolvedOutputDir,
    $InputPath, $RepVtasPath, $PxPath,
    $FacturaChanganPath, $NotaChanganPath, $MayorChanganPath,
    $FacturaPeugPath, $NotaPeugPath, $MayorPeugPath,
    $FacturaSzkPath, $NotaSzkPath, $MayorSzkPath,
    $FacturaTytPath, $NotaTytPath, $MayorTytPath
) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
Assert-NoVisibleExcelConflict -ConflictPaths $_conflictPaths
Register-OleMessageFilter

$excel = $null
$excelProcessId = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excelProcessId = Get-ExcelProcessId -Excel $excel
}
catch {
    Unregister-OleMessageFilter
    Release-Excel-Automation-Lock -Mutex $excelLock
    throw
}

Start-Sleep -Milliseconds 300
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false
$excel.EnableEvents = $false
$excel.AskToUpdateLinks = $false
try {
    $excel.Calculation = -4135
} catch {
}

$cancelMessage = $null
try {
    Assert-NotCancelled 'inicio'
    $timestamp = if ([string]::IsNullOrWhiteSpace($RunStamp)) {
        Get-Date -Format 'yyyyMMdd_HHmmss'
    } else {
        $RunStamp
    }

    $brandOrder = @('changan', 'peug', 'szk', 'tyt')
    if (-not [string]::IsNullOrWhiteSpace($BrandKey)) {
        $brandOrder = @($brandOrder | Where-Object { $_ -eq $BrandKey })
    }

    foreach ($templateKey in $brandOrder) {
        Assert-NotCancelled 'marcas'
        $brandStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $rowsDisplay = @($displayRows | Where-Object { $_.TemplateKey -eq $templateKey })
        $rowsRepVtas = @($repVtasRows | Where-Object { $_.TemplateKey -eq $templateKey })
        if ($rowsRepVtas.Count -eq 0 -and $rowsDisplay.Count -eq 0) {
            $brandStopwatch.Stop()
            continue
        }

        $rowsPosting = @($postingRows | Where-Object { $_.TemplateKey -eq $templateKey })
        Write-Output ("INFO|processing|{0}|rows={1}" -f $templateKey, $rowsRepVtas.Count)
        Write-Output ("INFO|validation_mode|{0}|{1}" -f $templateKey, $(if ($script:StrictValidationEnabled) { 'strict' } else { 'fast' }))

        $templateWorkbook = $null
        $openTemplateStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $templateWorkbook = Open-Workbook-WithRetry -Excel $excel -Path $templateConfigs[$templateKey].TemplatePath -ReadOnly $true
        $openTemplateStopwatch.Stop()
        Write-Output ("INFO|open_template_ms|{0}|{1}" -f $templateKey, $openTemplateStopwatch.ElapsedMilliseconds)
        try {
            Assert-NotCancelled 'base_plantilla'
            $templateSnapshot = Get-TemplateSnapshot -Workbook $templateWorkbook -TemplatePath $templateConfigs[$templateKey].TemplatePath -TemplateKey $templateKey
            $lookups = $templateSnapshot.Lookups
        }
        finally {
            # Se reutiliza esta plantilla para la validacion del mismo bloque.
        }

        $outputName = "servicios_{0}_{1}.xls" -f $templateKey, $timestamp
        $outputPath = Join-Path $resolvedOutputDir $outputName
        $copyTemplateStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        Copy-Item -LiteralPath $templateConfigs[$templateKey].TemplatePath -Destination $outputPath -Force
        $copyTemplateStopwatch.Stop()
        Write-Output ("INFO|copy_template_ms|{0}|{1}" -f $templateKey, $copyTemplateStopwatch.ElapsedMilliseconds)

        $openOutputStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $outputWorkbook = Open-Workbook-WithRetry -Excel $excel -Path $outputPath -ReadOnly $false
        $openOutputStopwatch.Stop()
        Write-Output ("INFO|open_output_ms|{0}|{1}" -f $templateKey, $openOutputStopwatch.ElapsedMilliseconds)
        $repSheet = $null
        $noteSheet = $null
        $repVtasSheet = $null
        $pxSheet = $null
        $mayorSheet = $null
        try {
            Assert-NotCancelled 'salida_plantilla'
            $repSheet = Get-Worksheet-Safe -Workbook $outputWorkbook -CandidateNames @('REP FACTURACION', 'REP FACTURACIÓN')
            $noteSheet = Get-Worksheet-Safe -Workbook $outputWorkbook -CandidateNames @('NOTA DE CREDITO')
            $repVtasSheet = Get-Worksheet-Safe -Workbook $outputWorkbook -CandidateNames @('REP VTAS')

            $fillInvoicesStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            $invoiceResult = Fill-Invoices -Worksheet $repSheet -Rows $rowsDisplay -Lookups $lookups
            $fillInvoicesStopwatch.Stop()
            Write-Output ("INFO|fill_invoices_ms|{0}|{1}" -f $templateKey, $fillInvoicesStopwatch.ElapsedMilliseconds)

            $fillNotesStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            $noteResult = Fill-Notes -Worksheet $noteSheet -Rows $rowsDisplay -Lookups $lookups
            $fillNotesStopwatch.Stop()
            Write-Output ("INFO|fill_notes_ms|{0}|{1}" -f $templateKey, $fillNotesStopwatch.ElapsedMilliseconds)

            $fillRepVtasStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            $repVtasResult = Fill-RepVtas -Worksheet $repVtasSheet -Rows $rowsRepVtas -Lookups $lookups
            $fillRepVtasStopwatch.Stop()
            Write-Output ("INFO|fill_repvtas_ms|{0}|{1}" -f $templateKey, $fillRepVtasStopwatch.ElapsedMilliseconds)

            $pxSheet = $null
            $pxRows = @()
            try { $pxSheet = Get-Worksheet-Safe -Workbook $outputWorkbook -CandidateNames @('PX') } catch {}
            if ($null -ne $pxSheet) {
                $fillPxStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
                $pxRows = Read-PxRows -PxPath $PxPath -BrandKey $templateKey -WorkingDirectory $stagingDirectory
                $pxResult = Write-PxRows-ToWorksheet -Worksheet $pxSheet -Rows $pxRows -BrandKey $templateKey
                $fillPxStopwatch.Stop()
                Write-Output ("INFO|fill_px_ms|{0}|{1}" -f $templateKey, $fillPxStopwatch.ElapsedMilliseconds)
                Write-Output ("INFO|px|{0}|rows={1}|bottom_capacity={2}" -f $templateKey, $pxResult.RowCount, $pxResult.BottomCapacity)
            }

            $mayorRows = @()
            $mayorResult = $null
            $mayorSheet = $null
            try { $mayorSheet = Get-Worksheet-Safe -Workbook $outputWorkbook -CandidateNames @('VENTAS', 'MAY VTAS') } catch {}
            if ($null -ne $mayorSheet) {
                $fillMayorStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
                $mayorPath = Resolve-MayorPathForBrand -BrandKey $templateKey -MayorPaths $mayorPaths
                if (-not [string]::IsNullOrWhiteSpace($mayorPath) -and (Test-Path -LiteralPath $mayorPath)) {
                    $mayorRows = Read-MayorRows -MayorPath $mayorPath -WorkingDirectory $stagingDirectory
                    $mayorFilter = Filter-MayorRowsForWorkbook -Rows $mayorRows
                    $mayorRows = @($mayorFilter.Rows)
                    if (@($mayorFilter.Removed).Count -gt 0) {
                        Write-Output ("INFO|mayor_px_adjustments_filtered|{0}|rows={1}" -f $templateKey, @($mayorFilter.Removed).Count)
                    }
                }

                $mayorResult = Write-MayorRows-ToWorksheet -Worksheet $mayorSheet -Rows $mayorRows -BrandKey $templateKey
                $fillMayorStopwatch.Stop()
                Write-Output ("INFO|fill_mayor_ms|{0}|{1}" -f $templateKey, $fillMayorStopwatch.ElapsedMilliseconds)
                if ($mayorRows.Count -gt 0) {
                    Write-Output ("INFO|mayor|{0}|rows={1}|sections={2}" -f $templateKey, $mayorResult.RowCount, $mayorResult.SectionCount)
                } else {
                    Write-Output ("WARN|mayor_missing|{0}|sheet_cleared" -f $templateKey)
                }
            }

            $costMetrics = Get-BrandCostMetrics -Rows $rowsRepVtas
            $periodDateValue = Get-Brand-PeriodDateValue -Rows $rowsRepVtas -MayorRows $mayorRows

            $precontVentasSheet = $null
            $precontCostos2Sheet = $null
            $precontCostosSheet = $null
            $costoSheet = $null
            $estadisticasSheet = $null
            $precontVentasCount = 0
            $precontCostos2Count = 0
            try { $precontVentasSheet = Get-Worksheet-Safe -Workbook $outputWorkbook -CandidateNames @('PrecontabilizacionVentas') } catch {}
            try { $precontCostos2Sheet = Get-Worksheet-Safe -Workbook $outputWorkbook -CandidateNames @('PrecontabilizacionCostos (2)') } catch {}
            try { $precontCostosSheet = Get-Worksheet-Safe -Workbook $outputWorkbook -CandidateNames @('PrecontabilizacionCostos') } catch {}
            try { $costoSheet = Get-Worksheet-Safe -Workbook $outputWorkbook -CandidateNames @('COSTO') } catch {}
            try { $estadisticasSheet = Get-Worksheet-Safe -Workbook $outputWorkbook -CandidateNames @('ESTADISTICAS') } catch {}

            try {
                if ($null -ne $precontVentasSheet) {
                    $precontVentasStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
                    $precontVentasPrototypes = @($templateSnapshot.PrecontVentasPrototypes)
                    if ($precontVentasPrototypes.Count -eq 0) {
                        $precontVentasPrototypes = @(Read-PrecontVentasPrototypes -Worksheet $precontVentasSheet)
                    }
                    $precontVentasRows = New-PrecontVentasGeneratedRows -Prototypes $precontVentasPrototypes -RowsPosting $rowsPosting -MayorRows $mayorRows -PxRows $pxRows
                    $precontVentasCount = Write-PrecontVentasGeneratedRows -Worksheet $precontVentasSheet -Rows $precontVentasRows
                    Refresh-WorksheetPivotTablesSafe -Worksheet $precontVentasSheet -Label 'PrecontabilizacionVentas'
                    Update-PrecontVentasControlFormulas -Worksheet $precontVentasSheet -Rows $precontVentasRows
                    Update-RepAndNoteControlFormulas -RepWorksheet $repSheet -NoteWorksheet $noteSheet -Rows $precontVentasRows
                    $precontVentasStopwatch.Stop()
                    Write-Output ("INFO|fill_precont_ventas_ms|{0}|{1}" -f $templateKey, $precontVentasStopwatch.ElapsedMilliseconds)
                    Write-Output ("INFO|precont_ventas|{0}|rows={1}" -f $templateKey, $precontVentasCount)
                }

                if ($null -ne $precontCostos2Sheet) {
                    $precontCostos2Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
                    $precontCostos2Prototypes = @($templateSnapshot.PrecontCostos2Prototypes)
                    if ($precontCostos2Prototypes.Count -eq 0) {
                        $precontCostos2Prototypes = @(Read-PrecontCostos2Prototypes -Worksheet $precontCostos2Sheet)
                    }
                    $precontCostos2Count = Write-PrecontCostos2GeneratedRows -Worksheet $precontCostos2Sheet -Prototypes $precontCostos2Prototypes -Metrics $costMetrics
                    Refresh-WorksheetPivotTablesSafe -Worksheet $precontCostos2Sheet -Label 'PrecontabilizacionCostos (2)'
                    $precontCostos2Stopwatch.Stop()
                    Write-Output ("INFO|fill_precont_costos2_ms|{0}|{1}" -f $templateKey, $precontCostos2Stopwatch.ElapsedMilliseconds)
                    Write-Output ("INFO|precont_costos2|{0}|rows={1}" -f $templateKey, $precontCostos2Count)
                }

                if ($null -ne $precontCostosSheet) {
                    Clear-PrecontCostosWorksheet -Worksheet $precontCostosSheet
                    Write-Output ("INFO|precont_costos_legacy_neutralized|{0}" -f $templateKey)
                }

                if ($null -ne $estadisticasSheet) {
                    $estadisticasStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
                    Write-EstadisticasGeneratedRows -Worksheet $estadisticasSheet -Metrics $costMetrics -PeriodDateValue $periodDateValue
                    $estadisticasStopwatch.Stop()
                    Write-Output ("INFO|fill_estadisticas_ms|{0}|{1}" -f $templateKey, $estadisticasStopwatch.ElapsedMilliseconds)
                    Write-Output ("INFO|estadisticas|{0}|repuestos={1}|lub={2}|subcont={3}|accesorios={4}" -f $templateKey, $costMetrics.CostoRepuestos, $costMetrics.CostoLubricantes, ([double]$costMetrics.Costo + [double]$costMetrics.CostoSubconNc), $costMetrics.CostoAccesorios)
                }

                if ($null -ne $costoSheet) {
                    $costoStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
                    Write-CostoGeneratedRows -Worksheet $costoSheet -Metrics $costMetrics -PeriodDateValue $periodDateValue
                    Refresh-WorksheetPivotTablesSafe -Worksheet $costoSheet -Label 'COSTO'
                    $costoStopwatch.Stop()
                    Write-Output ("INFO|fill_costo_ms|{0}|{1}" -f $templateKey, $costoStopwatch.ElapsedMilliseconds)
                    Write-Output ("INFO|costo|{0}|total={1}" -f $templateKey, ([double]$costMetrics.CostoRepuestos + [double]$costMetrics.CostoLubricantes + [double]$costMetrics.CostoAccesorios + [double]$costMetrics.Costo + [double]$costMetrics.CostoSubconNc))
                }
                if ($null -ne $pxSheet) {
                    Update-PxPeriodAnchor -Worksheet $pxSheet -PeriodDateValue $periodDateValue
                }
            } finally {
                if ($null -ne $estadisticasSheet) {
                    [void][Runtime.Interopservices.Marshal]::ReleaseComObject($estadisticasSheet)
                }
                if ($null -ne $costoSheet) {
                    [void][Runtime.Interopservices.Marshal]::ReleaseComObject($costoSheet)
                }
                if ($null -ne $precontCostosSheet) {
                    [void][Runtime.Interopservices.Marshal]::ReleaseComObject($precontCostosSheet)
                }
                if ($null -ne $precontCostos2Sheet) {
                    [void][Runtime.Interopservices.Marshal]::ReleaseComObject($precontCostos2Sheet)
                }
                if ($null -ne $precontVentasSheet) {
                    [void][Runtime.Interopservices.Marshal]::ReleaseComObject($precontVentasSheet)
                }
            }

            # Recalcular formulas de la plantilla despues de escribir datos
            $recalcStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            try {
                $outputWorkbook.Application.CalculateFull() | Out-Null
            } catch {
                try {
                    $outputWorkbook.Calculate() | Out-Null
                } catch {
                    Write-Output ("WARN|recalc_full_failed|{0}" -f $_.Exception.Message)
                }
            }
            $recalcStopwatch.Stop()
            Write-Output ("INFO|recalc_ms|{0}|{1}" -f $templateKey, $recalcStopwatch.ElapsedMilliseconds)

            if ($script:LayoutRestoreEnabled) {
                $layoutStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
                Restore-SelectedWorkbookLayoutFromTemplate `
                    -TemplateWorkbook $templateWorkbook `
                    -OutputWorkbook $outputWorkbook `
                    -SheetNames @(
                        'PrecontabilizacionVentas',
                        'PrecontabilizacionCostos (2)',
                        'COSTO'
                    ) `
                    -SheetRangeMap @{
                        'PrecontabilizacionVentas' = @{
                            StartRow = 1
                            EndRow = [Math]::Max(20, ([int]$precontVentasCount + 5))
                            LastColumn = 22
                        }
                        'PrecontabilizacionCostos (2)' = @{
                            StartRow = 1
                            EndRow = [Math]::Max(12, ([int]$precontCostos2Count + 5))
                            LastColumn = 10
                        }
                        'COSTO' = @{
                            StartRow = 1
                            EndRow = 12
                            LastColumn = 10
                        }
                    }
                $layoutStopwatch.Stop()
                Write-Output ("INFO|layout_ms|{0}|{1}" -f $templateKey, $layoutStopwatch.ElapsedMilliseconds)
            } else {
                Write-Output ("INFO|layout_skipped|{0}|mode=fast" -f $templateKey)
            }

            if ($script:StrictValidationEnabled) {
                $validateStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
                Write-Output ("INFO|validate_begin|{0}" -f $templateKey)
                Validate-Services-BrandOutput `
                    -OutputWorkbook $outputWorkbook `
                    -TemplateWorkbook $templateWorkbook `
                    -RepVtasResult $repVtasResult `
                    -InvoiceResult $invoiceResult `
                    -NoteResult $noteResult `
                    -MayorResult $mayorResult
                Write-Output ("INFO|validate_done|{0}" -f $templateKey)
                $validateStopwatch.Stop()
                Write-Output ("INFO|validate_ms|{0}|{1}" -f $templateKey, $validateStopwatch.ElapsedMilliseconds)
            }

            Assert-NotCancelled 'guardado_plantilla'
            $saveStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            Write-Output ("INFO|save_begin|{0}" -f $templateKey)
            Save-Workbook-WithRetry -Workbook $outputWorkbook -PathForError $outputPath
            Write-Output ("INFO|save_done|{0}" -f $templateKey)
            $saveStopwatch.Stop()
            Write-Output ("INFO|save_ms|{0}|{1}" -f $templateKey, $saveStopwatch.ElapsedMilliseconds)
            Write-Output ("OUTPUT|{0}|{1}" -f $outputName, $templateConfigs[$templateKey].Label)
            Write-Output ("INFO|{0}|invoice_fallbacks={1}|note_fallbacks={2}" -f $templateKey, [int]$invoiceResult.FallbackCount, [int]$noteResult.FallbackCount)
        }
        finally {
            if ($null -ne $mayorSheet) {
                [void][Runtime.Interopservices.Marshal]::ReleaseComObject($mayorSheet)
                $mayorSheet = $null
            }
            if ($null -ne $pxSheet) {
                [void][Runtime.Interopservices.Marshal]::ReleaseComObject($pxSheet)
                $pxSheet = $null
            }
            if ($null -ne $repVtasSheet) {
                [void][Runtime.Interopservices.Marshal]::ReleaseComObject($repVtasSheet)
                $repVtasSheet = $null
            }
            if ($null -ne $noteSheet) {
                [void][Runtime.Interopservices.Marshal]::ReleaseComObject($noteSheet)
                $noteSheet = $null
            }
            if ($null -ne $repSheet) {
                [void][Runtime.Interopservices.Marshal]::ReleaseComObject($repSheet)
                $repSheet = $null
            }
            $outputWorkbook.Close($true)
            [void][Runtime.Interopservices.Marshal]::ReleaseComObject($outputWorkbook)
            if ($null -ne $templateWorkbook) {
                $templateWorkbook.Close($false)
                [void][Runtime.Interopservices.Marshal]::ReleaseComObject($templateWorkbook)
                $templateWorkbook = $null
            }
            $brandStopwatch.Stop()
            Write-Output ("INFO|total_brand_ms|{0}|{1}" -f $templateKey, $brandStopwatch.ElapsedMilliseconds)
        }
    }
}
catch {
    if ($_.Exception.Message -like 'CANCELLED:*') {
        $cancelMessage = ($_.Exception.Message -replace '^CANCELLED:\s*', '').Trim()
    } else {
        throw
    }
}
finally {
    Stop-Excel-Application -Excel $excel
    if ($null -ne $excel) {
        try {
            [void][Runtime.Interopservices.Marshal]::FinalReleaseComObject($excel)
        } catch {
            try { [void][Runtime.Interopservices.Marshal]::ReleaseComObject($excel) } catch {}
        }
    }
    Unregister-OleMessageFilter
    Stop-WorkerExcelProcess -ProcessId $excelProcessId -WaitMilliseconds 500
    Stop-OrphanExcelProcesses -TimeoutSeconds 5
    Release-Excel-Automation-Lock -Mutex $excelLock
    if (-not [string]::IsNullOrWhiteSpace($stagingDirectory) -and (Test-Path -LiteralPath $stagingDirectory)) {
        Remove-Item -LiteralPath $stagingDirectory -Recurse -Force -ErrorAction SilentlyContinue
    }
    [GC]::Collect()
}

if ($null -ne $cancelMessage) {
    Write-Output ("CANCELLED|{0}" -f $cancelMessage)
    exit 130
}
