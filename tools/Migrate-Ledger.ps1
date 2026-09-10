[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$OriginalConfig,
    [Parameter(Mandatory=$true)][string]$Config,
    [Parameter(Mandatory=$true)][string]$Ledger,
    [Parameter(Mandatory=$true)][string]$Output,
    [string]$DataDir,
    [switch]$ConfirmOriginalDefinition
)
$ErrorActionPreference = 'Stop'
$params = @('-NoProfile','-ExecutionPolicy','Bypass','-File',(Join-Path (Split-Path -Parent $PSScriptRoot) 'src/ReaderDataViewer.ps1'),
    '-MigrateLedger','-OriginalConfig',$OriginalConfig,'-Config',$Config,'-BaselineLedger',$Ledger,'-Output',$Output)
if ($DataDir) { $params += @('-DataDir',$DataDir) }
if ($ConfirmOriginalDefinition) { $params += '-ConfirmOriginalDefinition' }
& powershell.exe @params
exit $LASTEXITCODE
