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
    [string]$FacturaPath = '',

    [Parameter(Mandatory = $false)]
    [string]$NotaPath = '',

    [Parameter(Mandatory = $false)]
    [string]$PxPath = '',

    [Parameter(Mandatory = $false)]
    [string]$RepVtasPath = '',

    [Parameter(Mandatory = $false)]
    [string]$MayorChanganPath = '',

    [Parameter(Mandatory = $false)]
    [string]$MayorPeugPath = '',

    [Parameter(Mandatory = $false)]
    [string]$MayorSzkPath = '',

    [Parameter(Mandatory = $false)]
    [string]$MayorTytPath = '',

    [Parameter(Mandatory = $false)]
    [string]$VentasPath = '',

    [Parameter(Mandatory = $false)]
    [string]$RiobambaPath = ''
)

$ErrorActionPreference = 'Stop'
$script:CancelFilePath = $CancelPath
$script:StrictValidationEnabled = ([string]$env:SERVICIOS_MARCAS_STRICT_VALIDATE -ne '0')
$script:AllowTemplateDataFallback = ([string]$env:SERVICIOS_MARCAS_ALLOW_TEMPLATE_FALLBACK -eq '1')

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

function Build-SourceRows-FromInputs {
    param(
        [string]$FacturaPath,
        [string]$NotaPath,
        [string]$RiobambaPath,
        [string]$BrandKey = ''
    )

    $rows = New-Object System.Collections.Generic.List[object]
    $appendRows = {
        param($items, [string]$docType, [string]$templateOverride = '')
        $index = $rows.Count + 1
        foreach ($item in $items) {
            $agencyRaw = $item.'Agencia :'
            if (-not $agencyRaw) { $agencyRaw = $item.'Agencia' }
            $agency = Normalize-Text $agencyRaw
            if ($agency -eq '') { continue }
            $templateKey = if (-not [string]::IsNullOrWhiteSpace($templateOverride)) { $templateOverride } else { Get-TemplateKey-FromAgency $agency }
            if ($templateKey -eq '') { continue }

            $serie = Normalize-Text ($item.'Serie')
            $documentRaw = if ($docType -eq 'DC') { $item.'Nota Cred.' } else { $item.'Factura' }
            if (-not $documentRaw) { $documentRaw = $item.'Factura' }
            if (-not $documentRaw) { $documentRaw = $item.'Nota Cred.' }
            $document = Normalize-Text $documentRaw
            $order = Normalize-Text ($item.'Orden')
            $cedulaRaw = $item.'Cedula'
            if (-not $cedulaRaw) { $cedulaRaw = $item.'C.I.' }
            $cedula = Normalize-Text $cedulaRaw
            $customer = Normalize-Text ($item.'Cliente')
            $affectedDocumentRaw = ''
            if ($docType -eq 'DC') {
                $affectedDocumentRaw = Normalize-Text ($item.'Factura')
            }
            $fechaText = Normalize-Text ($item.'Fecha')
            $dateFactValue = $null
            $dateNoteValue = $null
            if ($fechaText -ne '') {
                $parsed = Get-Date -Date $fechaText -ErrorAction SilentlyContinue
                if ($parsed) {
                    if ($docType -eq 'DC') {
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
            $total = [double]([decimal]($item.'Total'))
            $garExt = Normalize-Text ($item.'Gar.Ext.')
            $asiento = Normalize-Text ($item.'Asiento')
            $tv = Normalize-Text ($item.'T.V.')
            if (-not $subtotal) { $subtotal = 0.0 }
            if (-not $discount) { $discount = 0.0 }
            if (-not $iva12) { $iva12 = 0.0 }
            if (-not $iva15) { $iva15 = 0.0 }
            if (-not $total) { $total = 0.0 }

            $rows.Add([pscustomobject]@{
                RowIndex = $index
                TemplateKey = $templateKey
                Agency = $agency
                Center = ''
                Order = $order
                Advisor = ''
                Line = ''
                DocType = $docType
                Cedula = $cedula
                Customer = $customer
                DocumentRaw = $document
                DocumentTrim = $document
                Series = $serie
                FormaPago = $tv
                Authorization = ''
                DateFactValue = $dateFactValue
                DateNoteValue = $dateNoteValue
                NoteCredit = if ($docType -eq 'DC') { $total } else { $discount }
                TotalManoObra = $subtotal
                TotalSubcontratos = 0.0
                TotalInsumos = 0.0
                TotalServicio = 0.0
                TotalAccesorios = 0.0
                TotalRepuestos = 0.0
                Interes = 0.0
                Iva = $iva
                Iva12 = $iva12
                Iva15 = $iva15
                Total = $total
                Costo = 0.0
                CostoLubricantes = 0.0
                CostoAccesorios = 0.0
                CostoRepuestos = 0.0
                CostoPintura = 0.0
                CostoSubconNc = 0.0
                GarExt = $garExt
                AffectedDocumentTrim = Normalize-Text $affectedDocumentRaw
                AffectedDocumentRaw = $affectedDocumentRaw
                MotivoNc = ''
                ObservacionNc = ''
            }) | Out-Null
            $index++
        }
    }

    $facturas = Parse-TabFile -Path $FacturaPath -Kind 'factura'
    $notas = Parse-TabFile -Path $NotaPath -Kind 'nota'
    $riobamba = Parse-TabFile -Path $RiobambaPath -Kind 'factura'

    & $appendRows $facturas 'FA'
    & $appendRows $notas 'DC'
    $riobambaTemplate = $BrandKey
    if ([string]::IsNullOrWhiteSpace($riobambaTemplate)) {
        $riobambaTemplate = 'changan'
    }
    & $appendRows $riobamba 'FA' $riobambaTemplate

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

function Resolve-NodeBinary {
    $candidates = New-Object System.Collections.Generic.List[string]

    try {
        $nodeCommand = Get-Command node.exe -ErrorAction Stop
        if ($null -ne $nodeCommand -and -not [string]::IsNullOrWhiteSpace($nodeCommand.Source)) {
            $candidates.Add($nodeCommand.Source) | Out-Null
        }
    } catch {
    }

    try {
        $nodeCommand = Get-Command node -ErrorAction Stop
        if ($null -ne $nodeCommand -and -not [string]::IsNullOrWhiteSpace($nodeCommand.Source)) {
            $candidates.Add($nodeCommand.Source) | Out-Null
        }
    } catch {
    }

    $programFiles = if ([string]::IsNullOrWhiteSpace($env:ProgramFiles)) { 'C:\Program Files' } else { $env:ProgramFiles }
    $programFilesX86 = if ([string]::IsNullOrWhiteSpace(${env:ProgramFiles(x86)})) { 'C:\Program Files (x86)' } else { ${env:ProgramFiles(x86)} }

    foreach ($candidate in @(
        (Join-Path $programFiles 'nodejs\node.exe'),
        (Join-Path $programFilesX86 'nodejs\node.exe'),
        'C:\Program Files\nodejs\node.exe'
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

    throw 'No se encontro node.exe para leer el Excel fuente sin COM.'
}

function Read-SourceRows-FromFile {
    param(
        [string]$InputPath,
        [string]$WorkingDirectory
    )

    $nodeBinary = Resolve-NodeBinary
    $readerScript = Resolve-RequiredPath -Path (Join-Path $PSScriptRoot 'scripts\cxp\servicios_marcas\read_source.js') -Label 'read_source.js'
    $jsonPath = Join-Path $WorkingDirectory 'source_rows.json'
    if (Test-Path -LiteralPath $jsonPath) {
        Remove-Item -LiteralPath $jsonPath -Force -ErrorAction SilentlyContinue
    }

    $nodeOutput = & $nodeBinary $readerScript '--input' $InputPath '--output-json' $jsonPath 2>&1
    if ($LASTEXITCODE -ne 0) {
        $detail = (($nodeOutput | ForEach-Object { "$_" }) -join [Environment]::NewLine).Trim()
        if ($detail -eq '') {
            $detail = 'Node termino sin detalle.'
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

    $nodeBinary = Resolve-NodeBinary
    $readerScript = Resolve-RequiredPath -Path (Join-Path $PSScriptRoot 'scripts\cxp\servicios_marcas\read_px.js') -Label 'read_px.js'
    $jsonPath = Join-Path $WorkingDirectory 'px_rows.json'
    if (Test-Path -LiteralPath $jsonPath) {
        Remove-Item -LiteralPath $jsonPath -Force -ErrorAction SilentlyContinue
    }

    $nodeOutput = & $nodeBinary $readerScript '--input' $PxPath '--brand' $BrandKey '--output-json' $jsonPath 2>&1
    if ($LASTEXITCODE -ne 0) {
        $detail = (($nodeOutput | ForEach-Object { "$_" }) -join [Environment]::NewLine).Trim()
        if ($detail -eq '') { $detail = 'Node termino sin detalle.' }
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

function Read-MayorRows {
    param(
        [string]$MayorPath,
        [string]$WorkingDirectory
    )

    if ([string]::IsNullOrWhiteSpace($MayorPath) -or -not (Test-Path -LiteralPath $MayorPath)) {
        return @()
    }

    $nodeBinary = Resolve-NodeBinary
    $readerScript = Resolve-RequiredPath -Path (Join-Path $PSScriptRoot 'scripts\cxp\servicios_marcas\read_mayor_txt.js') -Label 'read_mayor_txt.js'
    $jsonPath = Join-Path $WorkingDirectory ("mayor_{0}.json" -f ([Guid]::NewGuid().ToString('N')))
    if (Test-Path -LiteralPath $jsonPath) {
        Remove-Item -LiteralPath $jsonPath -Force -ErrorAction SilentlyContinue
    }

    $nodeOutput = & $nodeBinary $readerScript '--input' $MayorPath '--output-json' $jsonPath 2>&1
    if ($LASTEXITCODE -ne 0) {
        $detail = (($nodeOutput | ForEach-Object { "$_" }) -join [Environment]::NewLine).Trim()
        if ($detail -eq '') { $detail = 'Node termino sin detalle.' }
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

function Assert-NoVisibleExcel {
    param([int]$TimeoutSeconds = 20)

    $deadline = [datetime]::UtcNow.AddSeconds([Math]::Max(1, $TimeoutSeconds))
    while ($true) {
        $procs = @(
            Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue |
                Where-Object { $_.MainWindowHandle -ne 0 }
        )
        if ($procs.Count -eq 0) { return }

        if ([datetime]::UtcNow -ge $deadline) {
            $pids = ($procs | Select-Object -ExpandProperty Id) -join ','
            throw "Hay procesos de Excel abiertos visibles (PID: $pids). Cierralos y vuelve a intentar."
        }

        Start-Sleep -Seconds 1
    }
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
    if ($text.Length -gt 1 -and $text.StartsWith("'")) {
        return $text.Substring(1)
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

    $errorTexts = @('#DIV/0!', '#N/A', '#NAME?', '#NULL!', '#NUM!', '#REF!', '#VALUE!')
    for ($row = $StartRow; $row -le $EndRow; $row++) {
        Assert-NotCancelled 'validacion_errores'
        for ($column = 1; $column -le $LastColumn; $column++) {
            $text = (Normalize-Text $Worksheet.Cells.Item($row, $column).Text).ToUpperInvariant()
            if ($text -in $errorTexts) {
                throw ("La hoja {0} contiene error de Excel en fila {1} columna {2}: {3}" -f $Label, $row, $column, $text)
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

    foreach ($entry in $Rows) {
        Assert-NotCancelled 'validacion_filas'
        $rowNumber = [int]$entry.RowNumber
        $values = $entry.Values

        foreach ($column in $TextColumns) {
            $expected = if ($values.ContainsKey($column)) { Normalize-ExcelLogicalText $values[$column] } else { '' }
            $actual = Get-ComparableCellText -Worksheet $OutputWorksheet -Row $rowNumber -Column $column
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

            $actualValue = Get-ComparableCellNumber -Worksheet $OutputWorksheet -Row $rowNumber -Column $column
            $actualText = Normalize-Text $OutputWorksheet.Cells.Item($rowNumber, $column).Text
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

            $actualDate = Get-Date-Value -Worksheet $OutputWorksheet -Row $rowNumber -Column $column
            if ($null -ne $actualDate -and $actualDate -ne '') {
                $actualDate = [double](Get-Date-Write-Value $actualDate)
            }
            if ($null -eq $expectedDate -and $null -eq $actualDate) {
                continue
            }

            if ($null -eq $expectedDate -or $null -eq $actualDate -or [Math]::Abs([double]$expectedDate - [double]$actualDate) -ge 0.01) {
                throw ("La hoja {0} tiene fecha incorrecta en fila {1} columna {2}. Esperado '{3}' y llego '{4}'." -f $Label, $rowNumber, $column, $expectedDate, $actualDate)
            }
        }

        foreach ($column in $FormatColumns) {
            $templateFormat = Normalize-Text $TemplateWorksheet.Cells.Item($rowNumber, $column).NumberFormat
            $outputFormat = Normalize-Text $OutputWorksheet.Cells.Item($rowNumber, $column).NumberFormat
            if ($templateFormat -ne $outputFormat) {
                throw ("La hoja {0} perdio formato en fila {1} columna {2}. Esperado '{3}' y llego '{4}'." -f $Label, $rowNumber, $column, $templateFormat, $outputFormat)
            }
        }
    }

    $nextRow = [int]$Rows[-1].RowNumber + 1
    if ((Normalize-Text $OutputWorksheet.Cells.Item($nextRow, $TrailingBlankColumn).Text) -ne '') {
        throw ("La hoja {0} conserva datos residuales despues de la ultima fila esperada ({1})." -f $Label, $Rows[-1].RowNumber)
    }

    Assert-NoExcelErrorsInRange -Worksheet $OutputWorksheet -StartRow ([int]$Rows[0].RowNumber) -EndRow ([int]$Rows[-1].RowNumber) -LastColumn $LastColumn -Label $Label
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
    for ($r = 0; $r -lt $rowCount; $r++) {
        $row = $Rows[$r]
        for ($c = 0; $c -lt $colCount; $c++) {
            $value = $null
            if ($row -is [System.Collections.ICollection] -and $c -lt $row.Count) {
                $value = $row[$c]
            }
            $Worksheet.Cells.Item($StartRow + $r, $StartColumn + $c).Value2 = $value
        }
    }
}

function Resolve-MayorPathForBrand {
    param(
        [string]$BrandKey,
        [hashtable]$MayorPaths,
        [string]$LegacyPath = ''
    )

    $key = (Normalize-Text $BrandKey).ToLowerInvariant()
    if ($MayorPaths.ContainsKey($key)) {
        $specificPath = Normalize-Text $MayorPaths[$key]
        if ($specificPath -ne '') {
            return $specificPath
        }
    }

    return (Normalize-Text $LegacyPath)
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
        }) | Out-Null
    }

    return $layouts.ToArray()
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
    Set-NumericCellSafe -Worksheet $Worksheet -Row $TargetRow -Column 11 -Value (To-Number $MayorRow.balance)
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
    $writtenCount = 0
    foreach ($layout in $layouts) {
        $knownAccounts[$layout.Account] = $true
        Clear-Worksheet-RangeContents -Worksheet $Worksheet -StartColumn 'A' -StartRow $layout.StartRow -EndColumn 'M' -EndRow $layout.EndRow
        $accountRows = if ($rowsByAccount.ContainsKey($layout.Account)) { @($rowsByAccount[$layout.Account]) } else { @() }
        $capacity = [int]($layout.EndRow - $layout.StartRow + 1)
        if ($accountRows.Count -gt $capacity) {
            throw ("El mayor para {0} excede la capacidad de la seccion {1} en {2}. Capacidad={3}, filas={4}" -f $BrandKey, $layout.Account, $Worksheet.Name, $capacity, $accountRows.Count)
        }

        $targetRow = [int]$layout.StartRow
        foreach ($accountRow in $accountRows) {
            Write-MayorDataRow -Worksheet $Worksheet -TargetRow $targetRow -MayorRow $accountRow
            $targetRow++
            $writtenCount++
        }
    }

    foreach ($account in $rowsByAccount.Keys) {
        if (-not $knownAccounts.ContainsKey($account)) {
            Write-Output ("WARN|mayor_account_unmapped|{0}|{1}|sheet={2}" -f $BrandKey, $account, $Worksheet.Name)
        }
    }

    return [pscustomobject]@{
        RowCount = $writtenCount
        SectionCount = $layouts.Count
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
    $anticipo = 0.0
    $neto = Round-Amount ($totalAmount - $anticipo)

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

            $null = $Worksheet.Cells.Item($targetRow, 1).Value2 = (Get-Excel-TextLiteral $agencyValue)
            $null = $Worksheet.Cells.Item($targetRow, 2).Value2 = "'" + $centerValue
            $null = $Worksheet.Cells.Item($targetRow, 3).Value2 = (Get-Excel-TextLiteral $orderValue)
            $null = $Worksheet.Cells.Item($targetRow, 4).Value2 = (Get-Excel-TextLiteral $advisorValue)
            $null = $Worksheet.Cells.Item($targetRow, 5).Value2 = (Get-Excel-TextLiteral $lineValue)
            $null = $Worksheet.Cells.Item($targetRow, 6).Value2 = "'" + $cedulaValue
            $null = $Worksheet.Cells.Item($targetRow, 7).Value2 = (Get-Excel-TextLiteral $customerValue)
            $null = $Worksheet.Cells.Item($targetRow, 8).Value2 = "'" + $documentRawValue
            $factDate = $factDateValue
            Set-DateCellValue -Worksheet $Worksheet -Row $targetRow -Column 9 -Value $factDate -Context 'REP_VTAS_FACT'
            $noteDate = $noteDateValue
            Set-DateCellValue -Worksheet $Worksheet -Row $targetRow -Column 10 -Value $noteDate -Context 'REP_VTAS_NC'
            Set-NumericCellSafe -Worksheet $Worksheet -Row $targetRow -Column 11 -Value $noteCreditValue
            Set-NumericCellSafe -Worksheet $Worksheet -Row $targetRow -Column 12 -Value $totalManoObraValue
            Set-NumericCellSafe -Worksheet $Worksheet -Row $targetRow -Column 13 -Value $totalSubcontratosValue
            Set-NumericCellSafe -Worksheet $Worksheet -Row $targetRow -Column 14 -Value $totalInsumosValue
            Set-NumericCellSafe -Worksheet $Worksheet -Row $targetRow -Column 15 -Value $totalServicioValue
            Set-NumericCellSafe -Worksheet $Worksheet -Row $targetRow -Column 16 -Value $totalAccesoriosValue
            Set-NumericCellSafe -Worksheet $Worksheet -Row $targetRow -Column 17 -Value $totalRepuestosValue
            Set-NumericCellSafe -Worksheet $Worksheet -Row $targetRow -Column 18 -Value $interesValue -BlankIfZero
            Set-NumericCellSafe -Worksheet $Worksheet -Row $targetRow -Column 19 -Value $ivaValue
            Set-NumericCellSafe -Worksheet $Worksheet -Row $targetRow -Column 20 -Value $totalValue
            Set-NumericCellSafe -Worksheet $Worksheet -Row $targetRow -Column 21 -Value $costoValue
            Set-NumericCellSafe -Worksheet $Worksheet -Row $targetRow -Column 22 -Value $costoLubricantesValue
            Set-NumericCellSafe -Worksheet $Worksheet -Row $targetRow -Column 23 -Value $costoAccesoriosValue
            Set-NumericCellSafe -Worksheet $Worksheet -Row $targetRow -Column 24 -Value $costoRepuestosValue
            Set-NumericCellSafe -Worksheet $Worksheet -Row $targetRow -Column 25 -Value $costoPinturaValue
            Set-NumericCellSafe -Worksheet $Worksheet -Row $targetRow -Column 26 -Value $costoSubconNcValue
            $null = $Worksheet.Cells.Item($targetRow, 27).Value2 = (Get-Excel-TextLiteral $garExtValue)

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

            $null = $Worksheet.Cells.Item($targetRow, 1).Value2 = (Get-Excel-TextLiteral $agencyValue)
            $null = $Worksheet.Cells.Item($targetRow, 2).Value2 = (Get-Excel-TextLiteral $seriesValue)
            $null = $Worksheet.Cells.Item($targetRow, 3).Value2 = $row.DocumentTrim
            $invoiceDate = Get-Date-Write-Value $row.DateFactValue
            Set-DateCellValue -Worksheet $Worksheet -Row $targetRow -Column 4 -Value $invoiceDate -Context 'REP_FACTURACION_FECHA'
            $null = $Worksheet.Cells.Item($targetRow, 5).Value2 = (Get-Excel-TextLiteral $orderValue)
            $null = $Worksheet.Cells.Item($targetRow, 6).Value2 = "'" + $cedulaValue
            $null = $Worksheet.Cells.Item($targetRow, 7).Value2 = (Get-Excel-TextLiteral $customerValue)
            $null = $Worksheet.Cells.Item($targetRow, 8).Value2 = [double]$subtotal
            $null = $Worksheet.Cells.Item($targetRow, 9).Value2 = [double]$discount
            $null = $Worksheet.Cells.Item($targetRow, 10).Value2 = [double]$netoConIva
            $null = $Worksheet.Cells.Item($targetRow, 11).Value2 = [double]$netoIva0Value
            $null = $Worksheet.Cells.Item($targetRow, 12).Value2 = [double]$iva15Amount
            $null = $Worksheet.Cells.Item($targetRow, 13).Value2 = [double]$iva12Amount
            $null = $Worksheet.Cells.Item($targetRow, 14).Value2 = [double]$interestAmount
            $null = $Worksheet.Cells.Item($targetRow, 15).Value2 = [double]$totalAmount
            $null = $Worksheet.Cells.Item($targetRow, 16).Value2 = $asientoValue
            $null = $Worksheet.Cells.Item($targetRow, 17).Value2 = (Get-Excel-TextLiteral $garExtValue)
            $null = $Worksheet.Cells.Item($targetRow, 18).Value2 = (Get-Excel-TextLiteral $tvValue)
            $null = $Worksheet.Cells.Item($targetRow, 19).Value2 = (Get-Excel-TextLiteral $markerValue)

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
                    12 = [double]$iva15Amount
                    13 = [double]$iva12Amount
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
            $invoiceValue = Get-PreferredSourceText -SourceRaw ([string]$row.AffectedDocumentRaw) -SourceNormalized (Normalize-Text $row.AffectedDocumentTrim) -LookupValue $(if ($null -ne $lookup) { $lookup.Invoice } else { '' })
            $orderValue = Get-PreferredSourceText -SourceRaw ([string]$row.OrderRaw) -SourceNormalized $(if ($orderKey -ne '') { $orderKey } else { Normalize-Text $row.Order }) -LookupValue $(if ($null -ne $lookup) { $lookup.Order } else { '' })
            $cedulaValue = Get-PreferredLookupText -LookupValue $(if ($null -ne $lookup) { $lookup.Cedula } else { '' }) -SourceRaw $row.CedulaRaw -SourceNormalized $row.Cedula
            $customerValue = Get-PreferredLookupText -LookupValue $(if ($null -ne $lookup) { $lookup.Customer } else { '' }) -SourceRaw $row.CustomerRaw -SourceNormalized $row.Customer
            $asientoValue = ''
            $garExtValue = Resolve-TemplateGarExt -LookupValue $(if ($null -ne $lookup) { $lookup.GarExt } else { '' }) -SourceRaw $row.GarExtRaw -SourceNormalized $row.GarExt -TemplateDefault $garExtDefault

            $rowNumber = $targetRow

            $null = $Worksheet.Cells.Item($targetRow, 1).Value2 = (Get-Excel-TextLiteral $agencyValue)
            $null = $Worksheet.Cells.Item($targetRow, 2).Value2 = $row.DocumentTrim
            $creditNoteDate = Get-Date-Write-Value $row.DateNoteValue
            Set-DateCellValue -Worksheet $Worksheet -Row $targetRow -Column 3 -Value $creditNoteDate -Context 'NOTA_CREDITO_FECHA'
            $null = $Worksheet.Cells.Item($targetRow, 4).Value2 = (Get-Excel-TextLiteral $kindValue)
            $null = $Worksheet.Cells.Item($targetRow, 5).Value2 = (Get-Excel-TextLiteral $seriesValue)
            $null = $Worksheet.Cells.Item($targetRow, 6).Value2 = (Get-Excel-TextLiteral $invoiceValue)
            $null = $Worksheet.Cells.Item($targetRow, 7).Value2 = (Get-Excel-TextLiteral $orderValue)
            $null = $Worksheet.Cells.Item($targetRow, 8).Value2 = "'" + $cedulaValue
            $null = $Worksheet.Cells.Item($targetRow, 9).Value2 = (Get-Excel-TextLiteral $customerValue)
            $null = $Worksheet.Cells.Item($targetRow, 10).Value2 = [double]$subtotal
            $null = $Worksheet.Cells.Item($targetRow, 11).Value2 = [double]$discount
            $null = $Worksheet.Cells.Item($targetRow, 12).Value2 = [double]$netoSinIva
            $null = $Worksheet.Cells.Item($targetRow, 13).Value2 = [double]$netoConIva
            $null = $Worksheet.Cells.Item($targetRow, 14).Value2 = [double]$iva15Amount
            $null = $Worksheet.Cells.Item($targetRow, 15).Value2 = [double]$iva12Amount
            $null = $Worksheet.Cells.Item($targetRow, 16).Value2 = [double]$interestAmount
            $null = $Worksheet.Cells.Item($targetRow, 17).Value2 = [double]$totalAmount
            $null = $Worksheet.Cells.Item($targetRow, 18).Value2 = [double]$anticipo
            $null = $Worksheet.Cells.Item($targetRow, 19).Value2 = [double]$neto
            $null = $Worksheet.Cells.Item($targetRow, 20).Value2 = (Get-Excel-TextLiteral $asientoValue)
            $null = $Worksheet.Cells.Item($targetRow, 21).Value2 = (Get-Excel-TextLiteral $garExtValue)

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
        [object]$NoteResult
    )

    $outputRepSheet = $null
    $outputNoteSheet = $null
    $outputRepVtasSheet = $null
    $templateRepSheet = $null
    $templateNoteSheet = $null
    $templateRepVtasSheet = $null
    try {
        $outputRepSheet = Get-Worksheet-Safe -Workbook $OutputWorkbook -CandidateNames @('REP FACTURACION', 'REP FACTURACIÃ“N')
        $outputNoteSheet = Get-Worksheet-Safe -Workbook $OutputWorkbook -CandidateNames @('NOTA DE CREDITO')
        $outputRepVtasSheet = Get-Worksheet-Safe -Workbook $OutputWorkbook -CandidateNames @('REP VTAS')

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
$customSource = @()
$customRequested = (-not [string]::IsNullOrWhiteSpace($FacturaPath)) -or (-not [string]::IsNullOrWhiteSpace($NotaPath)) -or (-not [string]::IsNullOrWhiteSpace($RiobambaPath))
if ($customRequested) {
    $customSource = Build-SourceRows-FromInputs -FacturaPath $FacturaPath -NotaPath $NotaPath -RiobambaPath $RiobambaPath -BrandKey $BrandKey
    if ($customSource.Count -gt 0) {
        Write-Output ("INFO|custom_source|rows={0}" -f $customSource.Count)
        $sourceRows = $customSource
    } else {
        Write-Output 'WARN|custom_source_empty|falling_back_to_excel'
    }
}

if ($sourceRows.Count -eq 0) {
    $sourceRows = Read-SourceRows-FromFile -InputPath $stagedInputPath -WorkingDirectory $stagingDirectory
}

if ($sourceRows.Count -eq 0) {
    throw 'Los archivos de entrada no contienen filas válidas para procesar (custom y fallback vacíos).'
}

$repVtasRows = Normalize-SourceRows -Rows $sourceRows
$postingRows = Normalize-SourceRows -Rows $sourceRows -ConsolidateInvoiceDocuments
    if ($repVtasRows.Count -eq 0 -and $postingRows.Count -eq 0) {
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
Stop-OrphanExcelProcesses -TimeoutSeconds 20
Assert-NoVisibleExcel -TimeoutSeconds 30
Register-OleMessageFilter

$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
}
catch {
    Unregister-OleMessageFilter
    Release-Excel-Automation-Lock -Mutex $excelLock
    throw
}

Start-Sleep -Milliseconds 1200
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
        $rowsRepVtas = @($repVtasRows | Where-Object { $_.TemplateKey -eq $templateKey })
        if ($rowsRepVtas.Count -eq 0) {
            continue
        }

        $rowsPosting = @($postingRows | Where-Object { $_.TemplateKey -eq $templateKey })
        Write-Output ("INFO|processing|{0}|rows={1}" -f $templateKey, $rowsRepVtas.Count)

        $templateWorkbook = $null
        $templateWorkbook = Open-Workbook-WithRetry -Excel $excel -Path $templateConfigs[$templateKey].TemplatePath -ReadOnly $true
        try {
            Assert-NotCancelled 'base_plantilla'
            $lookups = Read-TemplateLookups -Workbook $templateWorkbook
        }
        finally {
            # Se reutiliza esta plantilla para la validacion del mismo bloque.
        }

        $outputName = "servicios_{0}_{1}.xls" -f $templateKey, $timestamp
        $outputPath = Join-Path $resolvedOutputDir $outputName
        Copy-Item -LiteralPath $templateConfigs[$templateKey].TemplatePath -Destination $outputPath -Force

        $outputWorkbook = Open-Workbook-WithRetry -Excel $excel -Path $outputPath -ReadOnly $false
        try {
            Assert-NotCancelled 'salida_plantilla'
            $repSheet = Get-Worksheet-Safe -Workbook $outputWorkbook -CandidateNames @('REP FACTURACION', 'REP FACTURACIÓN')
            $noteSheet = Get-Worksheet-Safe -Workbook $outputWorkbook -CandidateNames @('NOTA DE CREDITO')
            $repVtasSheet = Get-Worksheet-Safe -Workbook $outputWorkbook -CandidateNames @('REP VTAS')

            $invoiceResult = Fill-Invoices -Worksheet $repSheet -Rows $rowsPosting -Lookups $lookups
            $noteResult = Fill-Notes -Worksheet $noteSheet -Rows $rowsPosting -Lookups $lookups
            $repVtasResult = Fill-RepVtas -Worksheet $repVtasSheet -Rows $rowsRepVtas -Lookups $lookups

            $pxSheet = $null
            try { $pxSheet = Get-Worksheet-Safe -Workbook $outputWorkbook -CandidateNames @('PX') } catch {}
            if ($null -ne $pxSheet) {
                $pxRows = Read-PxRows -PxPath $PxPath -BrandKey $templateKey -WorkingDirectory $stagingDirectory
                Clear-Worksheet-RangeContents -Worksheet $pxSheet -StartColumn 'A' -StartRow 6 -EndColumn 'R' -EndRow 65536
                Write-Rows-ToWorksheet -Worksheet $pxSheet -Rows $pxRows -StartRow 6 -StartColumn 1
            }

            $mayorSheet = $null
            try { $mayorSheet = Get-Worksheet-Safe -Workbook $outputWorkbook -CandidateNames @('VENTAS', 'MAY VTAS') } catch {}
            if ($null -ne $mayorSheet) {
                $mayorPath = Resolve-MayorPathForBrand -BrandKey $templateKey -MayorPaths $mayorPaths -LegacyPath $VentasPath
                if (-not [string]::IsNullOrWhiteSpace($mayorPath) -and (Test-Path -LiteralPath $mayorPath)) {
                    $mayorRows = Read-MayorRows -MayorPath $mayorPath -WorkingDirectory $stagingDirectory
                    $mayorResult = Write-MayorRows-ToWorksheet -Worksheet $mayorSheet -Rows $mayorRows -BrandKey $templateKey
                    Write-Output ("INFO|mayor|{0}|rows={1}|sections={2}" -f $templateKey, $mayorResult.RowCount, $mayorResult.SectionCount)
                } else {
                    Write-Output ("WARN|mayor_missing|{0}|template_sheet_preserved" -f $templateKey)
                }
            }

            # Recalcular formulas de la plantilla despues de escribir datos
            Recalculate-WorksheetSafe -Worksheet $repSheet -Label 'REP FACTURACION'
            Recalculate-WorksheetSafe -Worksheet $noteSheet -Label 'NOTA DE CREDITO'
            Recalculate-WorksheetSafe -Worksheet $repVtasSheet -Label 'REP VTAS'
            try {
                $outputWorkbook.ForceFullCalculation = $true
                $outputWorkbook.FullCalculationOnLoad = $true
                $outputWorkbook.Application.CalculateFullRebuild() | Out-Null
            } catch {
                Write-Output ("WARN|recalc_full_failed|{0}" -f $_.Exception.Message)
            }

            if ($script:StrictValidationEnabled) {
                Validate-Services-BrandOutput `
                    -OutputWorkbook $outputWorkbook `
                    -TemplateWorkbook $templateWorkbook `
                    -RepVtasResult $repVtasResult `
                    -InvoiceResult $invoiceResult `
                    -NoteResult $noteResult
            }

            Assert-NotCancelled 'guardado_plantilla'
            Save-Workbook-WithRetry -Workbook $outputWorkbook -PathForError $outputPath
            Write-Output ("OUTPUT|{0}|{1}" -f $outputName, $templateConfigs[$templateKey].Label)
            Write-Output ("INFO|{0}|invoice_fallbacks={1}|note_fallbacks={2}" -f $templateKey, [int]$invoiceResult.FallbackCount, [int]$noteResult.FallbackCount)
        }
        finally {
            $outputWorkbook.Close($true)
            [void][Runtime.Interopservices.Marshal]::ReleaseComObject($outputWorkbook)
            if ($null -ne $templateWorkbook) {
                $templateWorkbook.Close($false)
                [void][Runtime.Interopservices.Marshal]::ReleaseComObject($templateWorkbook)
                $templateWorkbook = $null
            }
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
    [void][Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
    Unregister-OleMessageFilter
    Stop-OrphanExcelProcesses -TimeoutSeconds 20
    Release-Excel-Automation-Lock -Mutex $excelLock
    if (-not [string]::IsNullOrWhiteSpace($stagingDirectory) -and (Test-Path -LiteralPath $stagingDirectory)) {
        Remove-Item -LiteralPath $stagingDirectory -Recurse -Force -ErrorAction SilentlyContinue
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

if ($null -ne $cancelMessage) {
    Write-Output ("CANCELLED|{0}" -f $cancelMessage)
    exit 130
}
