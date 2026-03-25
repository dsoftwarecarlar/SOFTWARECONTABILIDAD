param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$Args
)

$ErrorActionPreference = 'Stop'

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..\..')
$php = Join-Path $repoRoot '.tools\php82\php.exe'
$artisan = Join-Path $repoRoot 'laravel_app\artisan'

if (-not (Test-Path $php)) {
    throw "No existe PHP 8.2 local en $php"
}

if (-not (Test-Path $artisan)) {
    throw "No existe artisan en $artisan"
}

& $php $artisan @Args
exit $LASTEXITCODE

