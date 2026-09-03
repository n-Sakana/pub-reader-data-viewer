# ============================================================================
# build_worker.ps1 -- compile ZipWorker.cs into ZipWorker.exe with the in-box
#                     csc.exe. No SDK, no NuGet, nothing pre-installed.
#
# ASCII-only on purpose: Excel emits this file verbatim at run time, and
# Windows PowerShell 5.1 reads a BOM-less script in the ANSI code page, so a
# non-ASCII byte here could change how the script parses. Japanese explanation
# lives in README.md.
#
#   .\build_worker.ps1                       # build next to this script
#   .\build_worker.ps1 -Source X.cs -Out Y.exe
#   .\build_worker.ps1 -Launch -Token t1     # build then start the worker
#
# Used for two things:
#   1. producing prebuilt\ZipWorker.exe ahead of time ("prebuilt" variant)
#   2. being one of the three files Excel writes at run time ("emitted" variant)
# ============================================================================
[CmdletBinding()]
param(
  [string] $Source = "",
  [string] $Out    = "",
  [switch] $Launch,
  [string] $Token  = "zipbench",
  [string] $Mode   = "offscreen",
  [int]    $ParentPid = 0
)

$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrEmpty($Source)) { $Source = Join-Path $here 'ZipWorker.cs' }
if ([string]::IsNullOrEmpty($Out))    { $Out    = Join-Path $here 'ZipWorker.exe' }
$log = Join-Path (Split-Path -Parent $Out) 'build.log'

if (-not (Test-Path -LiteralPath $Source)) { throw "source not found: $Source" }

$csc = Join-Path $env:SystemRoot 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (-not (Test-Path -LiteralPath $csc)) {
  $csc = Join-Path $env:SystemRoot 'Microsoft.NET\Framework\v4.0.30319\csc.exe'
}
if (-not (Test-Path -LiteralPath $csc)) {
  'NO_CSC_FOUND' | Set-Content -LiteralPath $log -Encoding ascii
  throw 'csc.exe (in-box .NET Framework compiler) not found'
}

# /codepage:65001 matches how ZipWorker.cs is always written (UTF-8 with BOM).
$cscArgs = @(
  '/nologo', '/target:winexe', '/platform:anycpu', '/optimize+', '/codepage:65001',
  "/out:$Out",
  '/reference:System.dll', '/reference:System.Drawing.dll', '/reference:System.Windows.Forms.dll',
  $Source
)

$sw = [Diagnostics.Stopwatch]::StartNew()
$output = & $csc @cscArgs 2>&1
$sw.Stop()
$output | Out-File -LiteralPath $log -Encoding utf8

if (-not (Test-Path -LiteralPath $Out)) {
  Add-Content -LiteralPath $log -Value 'BUILD_FAILED'
  Write-Output "BUILD_FAILED  see $log"
  exit 1
}
Add-Content -LiteralPath $log -Value 'BUILD_OK'
Write-Output ("BUILD_OK  {0}  {1} bytes  {2} ms" -f $Out, (Get-Item -LiteralPath $Out).Length, $sw.ElapsedMilliseconds)

if ($Launch) {
  $wlog = Join-Path (Split-Path -Parent $Out) 'worker.log'
  Start-Process -FilePath $Out -WindowStyle Hidden -ArgumentList @(
    '--token', $Token, '--mode', $Mode, '--parentpid', "$ParentPid", '--log', $wlog
  ) | Out-Null
  Write-Output "LAUNCHED token=$Token mode=$Mode"
}
exit 0
