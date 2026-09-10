# Compatibility entry point. Distribution rules live in tools/Build.ps1.
[CmdletBinding()]
param(
    [string]$Root = '',
    [ValidateSet('both','folder','zip')][string]$Format = 'folder',
    [ValidateSet('sample','none')][string]$Data = 'sample',
    [string]$OutputRoot = '',
    [switch]$RunTests,
    [switch]$SkipValidation
)
$ErrorActionPreference = 'Stop'
if (-not $Root) { $Root = Split-Path -Parent $PSScriptRoot }
$parameters = @{Command='package'; Format=$Format; Data=$Data}
if ($OutputRoot) { $parameters.OutputRoot = $OutputRoot }
if ($RunTests) { $parameters.RunTests = $true }
if ($SkipValidation) { $parameters.SkipValidation = $true }
& (Join-Path $Root 'tools/Build.ps1') @parameters
exit $LASTEXITCODE
