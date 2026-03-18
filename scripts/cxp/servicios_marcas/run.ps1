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
    [string]$VentasPath = '',

    [Parameter(Mandatory = $false)]
    [string]$RiobambaPath = ''
)

$legacyScript = Join-Path (Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))) 'run_servicios_marcas.ps1'
if (-not (Test-Path -LiteralPath $legacyScript)) {
    throw "No se encontro el worker heredado: $legacyScript"
}

& $legacyScript `
    -InputPath $InputPath `
    -OutputDir $OutputDir `
    -TemplateDir $TemplateDir `
    -RunStamp $RunStamp `
    -CancelPath $CancelPath `
    -BrandKey $BrandKey `
    -FacturaPath $FacturaPath `
    -NotaPath $NotaPath `
    -PxPath $PxPath `
    -RepVtasPath $RepVtasPath `
    -VentasPath $VentasPath `
    -RiobambaPath $RiobambaPath

exit $LASTEXITCODE
