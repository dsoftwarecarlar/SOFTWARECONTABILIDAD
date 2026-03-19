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
    -FacturaChanganPath $FacturaChanganPath `
    -NotaChanganPath $NotaChanganPath `
    -PxPath $PxPath `
    -RepVtasPath $RepVtasPath `
    -MayorChanganPath $MayorChanganPath `
    -FacturaPeugPath $FacturaPeugPath `
    -NotaPeugPath $NotaPeugPath `
    -MayorPeugPath $MayorPeugPath `
    -FacturaSzkPath $FacturaSzkPath `
    -NotaSzkPath $NotaSzkPath `
    -MayorSzkPath $MayorSzkPath `
    -FacturaTytPath $FacturaTytPath `
    -NotaTytPath $NotaTytPath `
    -MayorTytPath $MayorTytPath

exit $LASTEXITCODE
