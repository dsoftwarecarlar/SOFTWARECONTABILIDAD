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
    [string]$CancelPath = ''
)

$ErrorActionPreference = 'Stop'
$script:CancelFilePath = $CancelPath

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

    if ($null -eq $Value -or $Value -eq '') {
        Set-CellValue -Worksheet $Worksheet -Row $Row -Column $Column -Value $null -Context $Context
        return
    }

    try {
        $Worksheet.Cells.Item($Row, $Column).Value = [datetime]::FromOADate([double]$Value)
    } catch {
        $valueType = if ($null -eq $Value) { 'null' } else { $Value.GetType().FullName }
        throw "No se pudo escribir fecha [$Context] row=$Row col=$Column type=$valueType value=$Value :: $($_.Exception.Message)"
    }
}

function Get-Template-Key {
    param(
        [object]$Agency,
        [object]$Series
    )

    $agencyText = (Normalize-Text $Agency).ToUpperInvariant()
    $seriesText = Normalize-Text $Series

    switch ($agencyText) {
        'CHANGAN' { return 'changan' }
        'PEUGEOT' { return 'peug' }
        'MATRIZ' { return 'tyt' }
        'SUZUKI AMBATO' { return 'szk' }
        'SUZUKI RIOBAMBA' {
            if ($seriesText -eq '007-606') {
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

function Read-TemplateLookups {
    param([object]$Workbook)

    Assert-NotCancelled 'lookup'
    $invoiceStore = New-LookupStore
    $noteStore = New-LookupStore
    $repVtasStore = New-LookupStore

    $invoiceSheet = Get-Worksheet-Safe -Workbook $Workbook -CandidateNames @('REP FACTURACION', 'REP FACTURACIÓN')
    $noteSheet = Get-Worksheet-Safe -Workbook $Workbook -CandidateNames @('NOTA DE CREDITO')

    $repVtasSheet = Get-Worksheet-Safe -Workbook $Workbook -CandidateNames @('REP VTAS')
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
    }
}

function Read-SourceRows {
    param([object]$Worksheet)

    Assert-NotCancelled 'lectura_fuente'
    $rows = New-Object System.Collections.Generic.List[object]
    $seen = @{}
    $lastRow = $Worksheet.UsedRange.Row + $Worksheet.UsedRange.Rows.Count - 1

    for ($row = 2; $row -le $lastRow; $row++) {
        Assert-NotCancelled 'lectura_fuente'
        $agency = Normalize-Text $Worksheet.Cells.Item($row, 1).Text
        if ($agency -eq '') {
            continue
        }

        $series = Normalize-Text $Worksheet.Cells.Item($row, 14).Text
        if ($series -eq '') {
            $series = Normalize-Text $Worksheet.Cells.Item($row, 13).Text
        }

        $templateKey = Get-Template-Key -Agency $agency -Series $series
        if ($null -eq $templateKey) {
            continue
        }

        $docType = (Normalize-Text $Worksheet.Cells.Item($row, 8).Text).ToUpperInvariant()
        if ($docType -notin @('FA', 'FC', 'DC', 'DE')) {
            continue
        }

        $documentRaw = Normalize-Text $Worksheet.Cells.Item($row, 12).Text
        if ($documentRaw -eq '') {
            continue
        }

        $order = Normalize-Text $Worksheet.Cells.Item($row, 3).Text
        $dateFactText = Normalize-Text $Worksheet.Cells.Item($row, 15).Text
        $dateNoteText = Normalize-Text $Worksheet.Cells.Item($row, 18).Text
        $affectedRaw = Normalize-Text $Worksheet.Cells.Item($row, 37).Text
        $totalText = Normalize-Text $Worksheet.Cells.Item($row, 28).Text
        $dedupeKey = @(
            $templateKey,
            $order.ToUpperInvariant(),
            $docType,
            $documentRaw,
            $dateFactText,
            $dateNoteText,
            $affectedRaw,
            $totalText
        ) -join '|'
        if ($seen.ContainsKey($dedupeKey)) {
            continue
        }
        $seen[$dedupeKey] = $true

        $rows.Add([pscustomobject]@{
            RowIndex = $row
            TemplateKey = $templateKey
            Agency = $agency
            Center = Normalize-Text $Worksheet.Cells.Item($row, 2).Text
            Order = $order
            Advisor = Normalize-Text $Worksheet.Cells.Item($row, 5).Text
            Line = Normalize-Text $Worksheet.Cells.Item($row, 7).Text
            DocType = $docType
            Cedula = Normalize-Text $Worksheet.Cells.Item($row, 9).Text
            Customer = Normalize-Text $Worksheet.Cells.Item($row, 10).Text
            DocumentRaw = $documentRaw
            DocumentTrim = Trim-Document $documentRaw
            Series = $series
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
            AffectedDocumentTrim = Trim-Document $affectedRaw
        })
    }

    return $rows
}

function Clear-OutputSheet {
    param(
        [object]$Worksheet,
        [int]$StartRow,
        [string]$LastColumn
    )

    $null = $Worksheet.Range("A${StartRow}:${LastColumn}65536").ClearContents()
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

function Get-Invoice-Asiento {
    param([pscustomobject]$Row)

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

function Fill-RepVtas {
    param(
        [object]$Worksheet,
        [object[]]$Rows,
        [hashtable]$Lookups
    )

    Clear-OutputSheet -Worksheet $Worksheet -StartRow 15 -LastColumn 'AL'
    $targetRow = 15
    $sortedRows = @(
        $Rows |
            Sort-Object @{
                Expression = {
                    $repLookup = Find-RepVtasEntry -Store $Lookups.RepVtas -DocKey $_.DocumentTrim -OrderKey $_.Order
                    if ($null -ne $repLookup) {
                        return [int]$repLookup.RowOrder
                    }

                    return 100000 + [int]$_.RowIndex
                }
            }
    )

    foreach ($row in $sortedRows) {
        try {
            Assert-NotCancelled 'rep_vtas'
            $lookup = if ($row.DocType -in @('FA', 'FC')) {
                Find-LookupEntry -Store $Lookups.Invoice -DocKey $row.DocumentTrim -OrderKey $row.Order
            } else {
                Find-LookupEntry -Store $Lookups.Note -DocKey $row.DocumentTrim -OrderKey (Strip-Order-Suffix $row.Order)
            }
            $repLookup = Find-RepVtasEntry -Store $Lookups.RepVtas -DocKey $row.DocumentTrim -OrderKey $row.Order
            $agencyValue = if ($null -ne $repLookup -and (Normalize-Text $repLookup.Agency) -ne '') { $repLookup.Agency } else { $row.Agency }
            $centerValue = if ($null -ne $repLookup -and (Normalize-Text $repLookup.Center) -ne '') { $repLookup.Center } else { $row.Center }
            $orderValue = if ($null -ne $repLookup -and (Normalize-Text $repLookup.Order) -ne '') { $repLookup.Order } else { $row.Order }
            $advisorValue = if ($null -ne $repLookup -and (Normalize-Text $repLookup.Advisor) -ne '') { $repLookup.Advisor } else { $row.Advisor }
            $lineValue = if ($null -ne $repLookup -and (Normalize-Text $repLookup.Line) -ne '') { $repLookup.Line } else { $row.Line }
            $documentRawValue = if ($null -ne $repLookup -and (Normalize-Text $repLookup.DocumentRaw) -ne '') { $repLookup.DocumentRaw } else { $row.DocumentRaw }
            $factDateValue = if ($null -ne $repLookup -and $null -ne $repLookup.DateFactValue) { $repLookup.DateFactValue } else { $row.DateFactValue }
            $noteDateValue = if ($null -ne $repLookup -and $null -ne $repLookup.DateNoteValue) { $repLookup.DateNoteValue } else { $row.DateNoteValue }
            $noteCreditValue = if ($null -ne $repLookup) { $repLookup.NoteCredit } else { $row.NoteCredit }
            $totalManoObraValue = if ($null -ne $repLookup) { $repLookup.TotalManoObra } else { $row.TotalManoObra }
            $totalSubcontratosValue = if ($null -ne $repLookup) { $repLookup.TotalSubcontratos } else { $row.TotalSubcontratos }
            $totalInsumosValue = if ($null -ne $repLookup) { $repLookup.TotalInsumos } else { $row.TotalInsumos }
            $totalServicioValue = if ($null -ne $repLookup) { $repLookup.TotalServicio } else { $row.TotalServicio }
            $totalAccesoriosValue = if ($null -ne $repLookup) { $repLookup.TotalAccesorios } else { $row.TotalAccesorios }
            $totalRepuestosValue = if ($null -ne $repLookup) { $repLookup.TotalRepuestos } else { $row.TotalRepuestos }
            $interesValue = if ($null -ne $repLookup) { $repLookup.Interes } else { $row.Interes }
            $ivaValue = if ($null -ne $repLookup) { $repLookup.Iva } else { $row.Iva }
            $totalValue = if ($null -ne $repLookup) { $repLookup.Total } else { $row.Total }
            $costoValue = if ($null -ne $repLookup) { $repLookup.Costo } else { $row.Costo }
            $costoLubricantesValue = if ($null -ne $repLookup) { $repLookup.CostoLubricantes } else { $row.CostoLubricantes }
            $costoAccesoriosValue = if ($null -ne $repLookup) { $repLookup.CostoAccesorios } else { $row.CostoAccesorios }
            $costoRepuestosValue = if ($null -ne $repLookup) { $repLookup.CostoRepuestos } else { $row.CostoRepuestos }
            $costoPinturaValue = if ($null -ne $repLookup) { $repLookup.CostoPintura } else { $row.CostoPintura }
            $costoSubconNcValue = if ($null -ne $repLookup) { $repLookup.CostoSubconNc } else { $row.CostoSubconNc }
            $garExtValue = if ($null -ne $repLookup -and (Normalize-Text $repLookup.GarExt) -ne '') { $repLookup.GarExt } else { $row.GarExt }

            $null = $Worksheet.Cells.Item($targetRow, 1).Value2 = $agencyValue
            $null = $Worksheet.Cells.Item($targetRow, 2).Value2 = "'" + $centerValue
            $null = $Worksheet.Cells.Item($targetRow, 3).Value2 = $orderValue
            $null = $Worksheet.Cells.Item($targetRow, 4).Value2 = $advisorValue
            $null = $Worksheet.Cells.Item($targetRow, 5).Value2 = $lineValue
            $null = $Worksheet.Cells.Item($targetRow, 6).Value2 = "'" + (Get-Display-Cedula -Lookup $repLookup -Fallback (Get-Display-Cedula -Lookup $lookup -Fallback $row.Cedula))
            $null = $Worksheet.Cells.Item($targetRow, 7).Value2 = (Get-Display-Customer -Lookup $repLookup -Fallback (Get-Display-Customer -Lookup $lookup -Fallback $row.Customer))
            $null = $Worksheet.Cells.Item($targetRow, 8).Value2 = "'" + $documentRawValue
            $factDate = Get-Date-Write-Value $factDateValue
            if ($null -eq $factDate) { $null = $Worksheet.Cells.Item($targetRow, 9).ClearContents() } else { $null = $Worksheet.Cells.Item($targetRow, 9).Value = [datetime]::FromOADate($factDate) }
            $noteDate = Get-Date-Write-Value $noteDateValue
            if ($null -eq $noteDate) { $null = $Worksheet.Cells.Item($targetRow, 10).ClearContents() } else { $null = $Worksheet.Cells.Item($targetRow, 10).Value = [datetime]::FromOADate($noteDate) }
            if ($null -ne $repLookup) {
                Set-TemplateNumericCell -Worksheet $Worksheet -Row $targetRow -Column 11 -TemplateText $repLookup.NoteCreditText -Value ([double]$noteCreditValue)
                Set-TemplateNumericCell -Worksheet $Worksheet -Row $targetRow -Column 12 -TemplateText $repLookup.TotalManoObraText -Value ([double]$totalManoObraValue)
                Set-TemplateNumericCell -Worksheet $Worksheet -Row $targetRow -Column 13 -TemplateText $repLookup.TotalSubcontratosText -Value ([double]$totalSubcontratosValue)
                Set-TemplateNumericCell -Worksheet $Worksheet -Row $targetRow -Column 14 -TemplateText $repLookup.TotalInsumosText -Value ([double]$totalInsumosValue)
                Set-TemplateNumericCell -Worksheet $Worksheet -Row $targetRow -Column 15 -TemplateText $repLookup.TotalServicioText -Value ([double]$totalServicioValue)
                Set-TemplateNumericCell -Worksheet $Worksheet -Row $targetRow -Column 16 -TemplateText $repLookup.TotalAccesoriosText -Value ([double]$totalAccesoriosValue)
                Set-TemplateNumericCell -Worksheet $Worksheet -Row $targetRow -Column 17 -TemplateText $repLookup.TotalRepuestosText -Value ([double]$totalRepuestosValue)
                Set-TemplateNumericCell -Worksheet $Worksheet -Row $targetRow -Column 18 -TemplateText $repLookup.InteresText -Value ([double]$interesValue)
                Set-TemplateNumericCell -Worksheet $Worksheet -Row $targetRow -Column 19 -TemplateText $repLookup.IvaText -Value ([double]$ivaValue)
                Set-TemplateNumericCell -Worksheet $Worksheet -Row $targetRow -Column 20 -TemplateText $repLookup.TotalText -Value ([double]$totalValue)
                Set-TemplateNumericCell -Worksheet $Worksheet -Row $targetRow -Column 21 -TemplateText $repLookup.CostoText -Value ([double]$costoValue)
                Set-TemplateNumericCell -Worksheet $Worksheet -Row $targetRow -Column 22 -TemplateText $repLookup.CostoLubricantesText -Value ([double]$costoLubricantesValue)
                Set-TemplateNumericCell -Worksheet $Worksheet -Row $targetRow -Column 23 -TemplateText $repLookup.CostoAccesoriosText -Value ([double]$costoAccesoriosValue)
                Set-TemplateNumericCell -Worksheet $Worksheet -Row $targetRow -Column 24 -TemplateText $repLookup.CostoRepuestosText -Value ([double]$costoRepuestosValue)
                Set-TemplateNumericCell -Worksheet $Worksheet -Row $targetRow -Column 25 -TemplateText $repLookup.CostoPinturaText -Value ([double]$costoPinturaValue)
                Set-TemplateNumericCell -Worksheet $Worksheet -Row $targetRow -Column 26 -TemplateText $repLookup.CostoSubconNcText -Value ([double]$costoSubconNcValue)
            } else {
                $null = $Worksheet.Cells.Item($targetRow, 11).Value2 = [double]$noteCreditValue
                $null = $Worksheet.Cells.Item($targetRow, 12).Value2 = [double]$totalManoObraValue
                $null = $Worksheet.Cells.Item($targetRow, 13).Value2 = [double]$totalSubcontratosValue
                $null = $Worksheet.Cells.Item($targetRow, 14).Value2 = [double]$totalInsumosValue
                $null = $Worksheet.Cells.Item($targetRow, 15).Value2 = [double]$totalServicioValue
                $null = $Worksheet.Cells.Item($targetRow, 16).Value2 = [double]$totalAccesoriosValue
                $null = $Worksheet.Cells.Item($targetRow, 17).Value2 = [double]$totalRepuestosValue
                if ((Round-Amount ([Math]::Abs($interesValue))) -eq 0) {
                    $null = $Worksheet.Cells.Item($targetRow, 18).ClearContents()
                } else {
                    $null = $Worksheet.Cells.Item($targetRow, 18).Value2 = [double]$interesValue
                }
                $null = $Worksheet.Cells.Item($targetRow, 19).Value2 = [double]$ivaValue
                $null = $Worksheet.Cells.Item($targetRow, 20).Value2 = [double]$totalValue
                $null = $Worksheet.Cells.Item($targetRow, 21).Value2 = [double]$costoValue
                $null = $Worksheet.Cells.Item($targetRow, 22).Value2 = [double]$costoLubricantesValue
                $null = $Worksheet.Cells.Item($targetRow, 23).Value2 = [double]$costoAccesoriosValue
                $null = $Worksheet.Cells.Item($targetRow, 24).Value2 = [double]$costoRepuestosValue
                $null = $Worksheet.Cells.Item($targetRow, 25).Value2 = [double]$costoPinturaValue
                $null = $Worksheet.Cells.Item($targetRow, 26).Value2 = [double]$costoSubconNcValue
            }
            $null = $Worksheet.Cells.Item($targetRow, 27).Value2 = $garExtValue
            $targetRow++
        } catch {
            throw "REP VTAS doc=$($row.DocumentRaw) orden=$($row.Order) filaFuente=$($row.RowIndex): $($_.Exception.Message)"
        }
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
    $writtenAggregateKeys = @{}
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
            if ($null -eq $lookup) {
                $fallbackCount++
            } elseif ((Normalize-Text $lookup.Order) -eq '') {
                $aggregateKey = @(
                    (Normalize-Text $row.Agency).ToUpperInvariant(),
                    (Normalize-Text $row.Series),
                    $row.DocumentTrim
                ) -join '|'
                if ($writtenAggregateKeys.ContainsKey($aggregateKey)) {
                    continue
                }
                $writtenAggregateKeys[$aggregateKey] = $true
            }

            $totalAmount = if ($null -ne $lookup) { Round-Amount $lookup.Total } else { Round-Amount ([Math]::Abs($row.Total)) }
            $ivaAmount = if ($null -ne $lookup) { Round-Amount $lookup.Iva } else { Round-Amount ([Math]::Abs($row.Iva)) }
            $interestAmount = if ($null -ne $lookup) { Round-Amount $lookup.Interest } else { Round-Amount ([Math]::Abs($row.Interes)) }
            $discount = if ($null -ne $lookup) { Round-Amount $lookup.Discount } else { 0 }
            $netoConIva = if ($null -ne $lookup) { Round-Amount $lookup.NetoConIva } else { Round-Amount ($totalAmount - $ivaAmount - $interestAmount) }
            $subtotal = if ($null -ne $lookup) { Round-Amount $lookup.Subtotal } else { Round-Amount ($netoConIva + $discount) }
            $agencyValue = if ($null -ne $lookup -and $lookup.Agency -ne '') { $lookup.Agency } else { $row.Agency }
            $seriesValue = if ($null -ne $lookup -and $lookup.Series -ne '') { $lookup.Series } else { $row.Series }
            $orderValue = if ($null -ne $lookup) { $lookup.Order } else { $row.Order }
            $netoIva0Value = if ($null -ne $lookup) { $lookup.NetoIva0 } else { 0 }
            $iva12Value = if ($null -ne $lookup) { $lookup.Iva12 } else { 0 }
            $asientoValue = if ($null -ne $lookup -and $lookup.Asiento -ne '') { $lookup.Asiento } else { Get-Invoice-Asiento -Row $row }
            $garExtValue = if ($null -ne $lookup) { $lookup.GarExt } else { '' }
            $tvValue = if ($null -ne $lookup -and $lookup.Tv -ne '') { $lookup.Tv } else { $row.Line }
            $markerValue = if ($null -ne $lookup -and $lookup.Marker -ne '') { $lookup.Marker } else { 'N' }

            $null = $Worksheet.Cells.Item($targetRow, 1).Value2 = $agencyValue
            $null = $Worksheet.Cells.Item($targetRow, 2).Value2 = $seriesValue
            $null = $Worksheet.Cells.Item($targetRow, 3).Value2 = $row.DocumentTrim
            $invoiceDate = Get-Date-Write-Value $row.DateFactValue
            if ($null -eq $invoiceDate) { $null = $Worksheet.Cells.Item($targetRow, 4).ClearContents() } else { $null = $Worksheet.Cells.Item($targetRow, 4).Value = [datetime]::FromOADate($invoiceDate) }
            $null = $Worksheet.Cells.Item($targetRow, 5).Value2 = $orderValue
            $null = $Worksheet.Cells.Item($targetRow, 6).Value2 = "'" + (Get-Display-Cedula -Lookup $lookup -Fallback $row.Cedula)
            $null = $Worksheet.Cells.Item($targetRow, 7).Value2 = (Get-Display-Customer -Lookup $lookup -Fallback $row.Customer)
            $null = $Worksheet.Cells.Item($targetRow, 8).Value2 = [double]$subtotal
            $null = $Worksheet.Cells.Item($targetRow, 9).Value2 = [double]$discount
            $null = $Worksheet.Cells.Item($targetRow, 10).Value2 = [double]$netoConIva
            $null = $Worksheet.Cells.Item($targetRow, 11).Value2 = [double]$netoIva0Value
            $null = $Worksheet.Cells.Item($targetRow, 12).Value2 = [double]$iva12Value
            if ($null -ne $lookup -and (Normalize-Text $lookup.IvaText) -eq '') {
                $null = $Worksheet.Cells.Item($targetRow, 13).ClearContents()
            } else {
                $null = $Worksheet.Cells.Item($targetRow, 13).Value2 = [double]$ivaAmount
            }
            $null = $Worksheet.Cells.Item($targetRow, 14).Value2 = [double]$interestAmount
            $null = $Worksheet.Cells.Item($targetRow, 15).Value2 = [double]$totalAmount
            $null = $Worksheet.Cells.Item($targetRow, 16).Value2 = $asientoValue
            $null = $Worksheet.Cells.Item($targetRow, 17).Value2 = $garExtValue
            $null = $Worksheet.Cells.Item($targetRow, 18).Value2 = $tvValue
            $null = $Worksheet.Cells.Item($targetRow, 19).Value2 = $markerValue
            $targetRow++
        } catch {
            throw "FACTURA doc=$($row.DocumentTrim) orden=$($row.Order) filaFuente=$($row.RowIndex): $($_.Exception.Message)"
        }
    }

    return $fallbackCount
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
            if ($null -eq $lookup) {
                $fallbackCount++
            }

            $totalAmount = if ($null -ne $lookup) { Round-Amount $lookup.Total } else { Round-Amount ([Math]::Abs($row.Total)) }
            $ivaAmount = if ($null -ne $lookup) { Round-Amount $lookup.Iva } else { Round-Amount ([Math]::Abs($row.Iva)) }
            $interestAmount = if ($null -ne $lookup) { Round-Amount $lookup.Interest } else { Round-Amount ([Math]::Abs($row.Interes)) }
            $discount = if ($null -ne $lookup) { Round-Amount $lookup.Discount } else { 0 }
            $netoSinIva = if ($null -ne $lookup) { Round-Amount $lookup.NetoSinIva } else { 0 }
            $netoConIva = if ($null -ne $lookup) { Round-Amount $lookup.NetoConIva } else { Round-Amount ($totalAmount - $ivaAmount - $interestAmount) }
            $subtotal = if ($null -ne $lookup) { Round-Amount $lookup.Subtotal } else { Round-Amount ($netoConIva + $discount) }
            $anticipo = if ($null -ne $lookup) { Round-Amount $lookup.Anticipo } else { 0 }
            $neto = if ($null -ne $lookup) { Round-Amount $lookup.Neto } else { Round-Amount ($totalAmount - $anticipo) }
            $agencyValue = if ($null -ne $lookup -and $lookup.Agency -ne '') { $lookup.Agency } else { $row.Agency }
            $kindValue = if ($null -ne $lookup -and $lookup.Kind -ne '') {
                $lookup.Kind
            } elseif ($row.DocType -eq 'DE') {
                'CON'
            } else {
                'CRE'
            }
            $seriesValue = if ($null -ne $lookup -and $lookup.Series -ne '') { $lookup.Series } else { $row.Series }
            $invoiceValue = if ($null -ne $lookup -and $lookup.Invoice -ne '') { $lookup.Invoice } else { $row.AffectedDocumentTrim }
            $orderValue = if ($null -ne $lookup -and $lookup.Order -ne '') { $lookup.Order } else { $orderKey }
            $iva12Value = if ($null -ne $lookup) { $lookup.Iva12 } else { 0 }
            $asientoValue = if ($null -ne $lookup) { $lookup.Asiento } else { '' }
            $garExtValue = if ($null -ne $lookup -and $lookup.GarExt -ne '') { $lookup.GarExt } else { 'N' }

            $null = $Worksheet.Cells.Item($targetRow, 1).Value2 = $agencyValue
            $null = $Worksheet.Cells.Item($targetRow, 2).Value2 = $row.DocumentTrim
            $creditNoteDate = Get-Date-Write-Value $row.DateNoteValue
            if ($null -eq $creditNoteDate) { $null = $Worksheet.Cells.Item($targetRow, 3).ClearContents() } else { $null = $Worksheet.Cells.Item($targetRow, 3).Value = [datetime]::FromOADate($creditNoteDate) }
            $null = $Worksheet.Cells.Item($targetRow, 4).Value2 = $kindValue
            $null = $Worksheet.Cells.Item($targetRow, 5).Value2 = $seriesValue
            $null = $Worksheet.Cells.Item($targetRow, 6).Value2 = $invoiceValue
            $null = $Worksheet.Cells.Item($targetRow, 7).Value2 = $orderValue
            $null = $Worksheet.Cells.Item($targetRow, 8).Value2 = "'" + (Get-Display-Cedula -Lookup $lookup -Fallback $row.Cedula)
            $null = $Worksheet.Cells.Item($targetRow, 9).Value2 = (Get-Display-Customer -Lookup $lookup -Fallback $row.Customer)
            $null = $Worksheet.Cells.Item($targetRow, 10).Value2 = [double]$subtotal
            $null = $Worksheet.Cells.Item($targetRow, 11).Value2 = [double]$discount
            $null = $Worksheet.Cells.Item($targetRow, 12).Value2 = [double]$netoSinIva
            $null = $Worksheet.Cells.Item($targetRow, 13).Value2 = [double]$netoConIva
            $null = $Worksheet.Cells.Item($targetRow, 14).Value2 = [double]$ivaAmount
            $null = $Worksheet.Cells.Item($targetRow, 15).Value2 = [double]$iva12Value
            $null = $Worksheet.Cells.Item($targetRow, 16).Value2 = [double]$interestAmount
            $null = $Worksheet.Cells.Item($targetRow, 17).Value2 = [double]$totalAmount
            $null = $Worksheet.Cells.Item($targetRow, 18).Value2 = [double]$anticipo
            $null = $Worksheet.Cells.Item($targetRow, 19).Value2 = [double]$neto
            $null = $Worksheet.Cells.Item($targetRow, 20).Value2 = $asientoValue
            $null = $Worksheet.Cells.Item($targetRow, 21).Value2 = $garExtValue
            $targetRow++
        } catch {
            throw "NOTA doc=$($row.DocumentTrim) orden=$($row.Order) filaFuente=$($row.RowIndex): $($_.Exception.Message)"
        }
    }

    return $fallbackCount
}

$resolvedInputPath = (Resolve-Path $InputPath).Path
$resolvedOutputDir = (Resolve-Path $OutputDir).Path
$resolvedTemplateDir = (Resolve-Path $TemplateDir).Path

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

$excel = New-Object -ComObject Excel.Application
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
    $sourceWorkbook = $excel.Workbooks.Open($resolvedInputPath, 0, $false)
    try {
        Assert-NotCancelled 'inicio'
        $sourceRows = Read-SourceRows -Worksheet $sourceWorkbook.Worksheets.Item(1)
    }
    finally {
        $sourceWorkbook.Close($false)
        [void][Runtime.Interopservices.Marshal]::ReleaseComObject($sourceWorkbook)
    }

    $timestamp = if ([string]::IsNullOrWhiteSpace($RunStamp)) {
        Get-Date -Format 'yyyyMMdd_HHmmss'
    } else {
        $RunStamp
    }

    foreach ($templateKey in @('changan', 'peug', 'szk', 'tyt')) {
        Assert-NotCancelled 'marcas'
        $rows = @($sourceRows | Where-Object { $_.TemplateKey -eq $templateKey })
        if ($rows.Count -eq 0) {
            continue
        }

        Write-Output ("INFO|processing|{0}|rows={1}" -f $templateKey, $rows.Count)

        $baseWorkbook = $excel.Workbooks.Open($templateConfigs[$templateKey].TemplatePath, 0, $false)
        try {
            Assert-NotCancelled 'base_plantilla'
            $lookups = Read-TemplateLookups -Workbook $baseWorkbook
        }
        finally {
            $baseWorkbook.Close($false)
            [void][Runtime.Interopservices.Marshal]::ReleaseComObject($baseWorkbook)
        }

        $outputName = "servicios_{0}_{1}.xls" -f $templateKey, $timestamp
        $outputPath = Join-Path $resolvedOutputDir $outputName
        Copy-Item -LiteralPath $templateConfigs[$templateKey].TemplatePath -Destination $outputPath -Force

        $outputWorkbook = $excel.Workbooks.Open($outputPath, 0, $false)
        try {
            Assert-NotCancelled 'salida_plantilla'
            $repSheet = Get-Worksheet-Safe -Workbook $outputWorkbook -CandidateNames @('REP FACTURACION', 'REP FACTURACIÓN')
            $noteSheet = Get-Worksheet-Safe -Workbook $outputWorkbook -CandidateNames @('NOTA DE CREDITO')
            $repVtasSheet = Get-Worksheet-Safe -Workbook $outputWorkbook -CandidateNames @('REP VTAS')

            $invoiceFallbacks = Fill-Invoices -Worksheet $repSheet -Rows $rows -Lookups $lookups
            $noteFallbacks = Fill-Notes -Worksheet $noteSheet -Rows $rows -Lookups $lookups
            Fill-RepVtas -Worksheet $repVtasSheet -Rows $rows -Lookups $lookups

            Assert-NotCancelled 'guardado_plantilla'
            $outputWorkbook.Save()
            Write-Output ("OUTPUT|{0}|{1}" -f $outputName, $templateConfigs[$templateKey].Label)
            Write-Output ("INFO|{0}|invoice_fallbacks={1}|note_fallbacks={2}" -f $templateKey, [int]$invoiceFallbacks, [int]$noteFallbacks)
        }
        finally {
            $outputWorkbook.Close($true)
            [void][Runtime.Interopservices.Marshal]::ReleaseComObject($outputWorkbook)
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
    $excel.Quit()
    [void][Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

if ($null -ne $cancelMessage) {
    Write-Output ("CANCELLED|{0}" -f $cancelMessage)
    exit 130
}
