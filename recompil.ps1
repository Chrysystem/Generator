# Wrapper vers recompile.ps1
param(
    [Parameter(ValueFromRemainingArguments=$true)]
    [string[]]$ArgsPassThru
)

$script = Join-Path -Path $PSScriptRoot -ChildPath 'recompile.ps1'
if (-not (Test-Path $script)) {
    Write-Host "ERREUR: recompile.ps1 introuvable dans $PSScriptRoot" -ForegroundColor Red
    exit 1
}

& $script @ArgsPassThru
