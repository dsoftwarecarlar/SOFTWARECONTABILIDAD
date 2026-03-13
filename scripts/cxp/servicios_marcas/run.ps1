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

$legacyScript = Join-Path (Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))) 'run_servicios_marcas.ps1'
if (-not (Test-Path -LiteralPath $legacyScript)) {
    throw "No se encontro el worker heredado: $legacyScript"
}

& $legacyScript `
    -InputPath $InputPath `
    -OutputDir $OutputDir `
    -TemplateDir $TemplateDir `
    -RunStamp $RunStamp `
    -CancelPath $CancelPath

exit $LASTEXITCODE
