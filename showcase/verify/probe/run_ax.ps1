# run_ax.ps1 -- VERIFICATION ONLY. Watchdog for the ActiveX probe.
[CmdletBinding()]
param(
    [int] $TimeoutSec = 70,
    [switch] $Visible,
    [switch] $SaveFirst,
    [switch] $ClearExdCache,
    [string] $Tag = 'ax'
)
$ErrorActionPreference = 'Stop'

$here    = Split-Path -Parent $MyInvocation.MyCommand.Path
$bas     = Join-Path $here 'modPbAx.bas'
$child   = Join-Path $here 'ax_child.ps1'
$res     = Join-Path $here "ax-result-$Tag.txt"
$pidFile = Join-Path $here "ax-owned-$Tag.pid"
$log     = Join-Path $here "ax-child-$Tag.log"
foreach ($f in @($res, $pidFile, $log)) { if (Test-Path $f) { Remove-Item $f -Force } }

function Read-Shared([string]$p) {
    if (-not (Test-Path $p)) { return '' }
    try {
        $fs = [System.IO.File]::Open($p, 'Open', 'Read', 'ReadWrite')
        $sr = New-Object System.IO.StreamReader($fs)
        $t  = $sr.ReadToEnd(); $sr.Close(); $fs.Close(); return $t
    } catch { return '' }
}

if ($ClearExdCache) {
    # MSForms.exd is a REGENERATED cache under %TEMP%. A stale one is the known
    # cause of "cannot insert object" after an Office update. Backed up, not
    # thrown away, and nothing outside Temp is touched.
    $bak = Join-Path $here 'exd-backup'
    if (-not (Test-Path $bak)) { New-Item -ItemType Directory -Path $bak | Out-Null }
    Get-ChildItem -Path $env:TEMP, "$env:LOCALAPPDATA\Temp" -Recurse -Filter *.exd -ErrorAction SilentlyContinue |
        Sort-Object FullName -Unique | ForEach-Object {
            Copy-Item $_.FullName (Join-Path $bak $_.Name) -Force -ErrorAction SilentlyContinue
            Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue
            Write-Host "cleared exd cache: $($_.FullName)"
        }
}

$childArgs = @('-ExecutionPolicy','Bypass','-NoProfile','-File',"`"$child`"",
               '-Bas',"`"$bas`"",'-Res',"`"$res`"",'-PidFile',"`"$pidFile`"",
               '-WaitSec',[string]($TimeoutSec - 20))
if ($Visible)   { $childArgs += '-Visible' }
if ($SaveFirst) { $childArgs += '-SaveFirst' }
$p = Start-Process -FilePath 'powershell.exe' -ArgumentList $childArgs -PassThru `
        -RedirectStandardOutput $log -RedirectStandardError "$log.err" -WindowStyle Hidden

$deadline = (Get-Date).AddSeconds($TimeoutSec)
while (-not $p.HasExited -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 400 }
if (-not $p.HasExited) {
    Write-Host 'DEADLINE: killing what the child owns'
    foreach ($id in ((Read-Shared $pidFile).Trim() -split ',' | Where-Object { $_ })) {
        if (Get-Process -Id ([int]$id) -ErrorAction SilentlyContinue) { Stop-Process -Id ([int]$id) -Force }
    }
    Start-Sleep -Milliseconds 400
    if (-not $p.HasExited) { Stop-Process -Id $p.Id -Force }
}
foreach ($id in ((Read-Shared $pidFile).Trim() -split ',' | Where-Object { $_ })) {
    if (Get-Process -Id ([int]$id) -ErrorAction SilentlyContinue) { Stop-Process -Id ([int]$id) -Force }
}
Write-Host "--- child log ($Tag) ---"
if (Test-Path $log) { Get-Content $log | ForEach-Object { "  $_" } }
if ((Test-Path "$log.err") -and (Get-Item "$log.err").Length -gt 0) { Get-Content "$log.err" | ForEach-Object { "  ERR $_" } }
Write-Host "--- result ($Tag) ---"
$t = Read-Shared $res
if ($t) { $t -split "`r?`n" | Where-Object { $_ } | ForEach-Object { "  $_" } } else { '  (empty)' }
