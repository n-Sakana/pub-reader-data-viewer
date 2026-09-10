param([string]$Root='', [string]$Config='', [string]$DataDir='', [string]$Evidence='', [string]$ActualConfig='')
$ErrorActionPreference='Stop'
& (Join-Path $PSScriptRoot 'Test-PaymentBasis.ps1') @PSBoundParameters
