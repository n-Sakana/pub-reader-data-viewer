# run_probe.ps1 -- VERIFICATION ONLY, not part of the deliverable.
#
# The watchdog. Every COM call happens in probe_child.ps1, which starts its
# Excel INVISIBLE and compile-checks the module before anything is shown. A VBA
# compile error is a modal, and a modal that nobody can see is a hang -- so this
# process never makes a COM call itself, and kills what the child owns on the
# deadline. Nothing is ever left on the operator's screen and the operator is
# never asked to click anything.
[CmdletBinding()]
param(
    [int] $TimeoutSec = 240,
    [string] $Skip = '',
    [string] $Out = ''
)
$ErrorActionPreference = 'Stop'

$here    = Split-Path -Parent $MyInvocation.MyCommand.Path
$bas     = Join-Path $here 'modPbProbe.bas'
$child   = Join-Path $here 'probe_child.ps1'
$res     = if ($Out) { $Out } else { Join-Path $here 'probe-result.txt' }
$pidFile = Join-Path $here 'probe-owned.pid'
$log     = Join-Path $here 'probe-child.log'
foreach ($f in @($res, $pidFile, $log)) { if (Test-Path $f) { Remove-Item $f -Force } }

function Read-Shared([string]$p) {
    if (-not (Test-Path $p)) { return '' }
    try {
        $fs = [System.IO.File]::Open($p, 'Open', 'Read', 'ReadWrite')
        $sr = New-Object System.IO.StreamReader($fs)
        $t  = $sr.ReadToEnd(); $sr.Close(); $fs.Close(); return $t
    } catch { return '' }
}

$childArgs = @('-ExecutionPolicy','Bypass','-NoProfile','-File',"`"$child`"",
               '-Bas',"`"$bas`"",'-Res',"`"$res`"",'-PidFile',"`"$pidFile`"",
               '-WaitSec',[string]($TimeoutSec - 20))
if ($Skip) { $childArgs += @('-Skip', "`"$Skip`"") }
$p = Start-Process -FilePath 'powershell.exe' -ArgumentList $childArgs -PassThru `
        -RedirectStandardOutput $log -RedirectStandardError "$log.err" -WindowStyle Hidden

$deadline = (Get-Date).AddSeconds($TimeoutSec)
$last = ''
while (-not $p.HasExited -and (Get-Date) -lt $deadline) {
    Start-Sleep -Milliseconds 500
    $txt = Read-Shared $res
    if ($txt) {
        $lines = @($txt -split "`r?`n" | Where-Object { $_ -ne '' })
        if ($lines.Count -gt 0 -and $lines[-1] -ne $last) {
            $last = $lines[-1]
            Write-Host "  ... $last"
        }
    }
}

if (-not $p.HasExited) {
    Write-Host 'DEADLINE: the child is stuck (a modal, or a UIA call that never returns)'
    $owned = (Read-Shared $pidFile).Trim()
    foreach ($id in ($owned -split ',' | Where-Object { $_ })) {
        if (Get-Process -Id ([int]$id) -ErrorAction SilentlyContinue) {
            Write-Host "  killing owned excel pid $id"
            Stop-Process -Id ([int]$id) -Force
        }
    }
    Start-Sleep -Milliseconds 500
    if (-not $p.HasExited) { Stop-Process -Id $p.Id -Force }
}

Write-Host '--- child log ---'
if (Test-Path $log) { Get-Content $log | ForEach-Object { "  $_" } }
if ((Test-Path "$log.err") -and (Get-Item "$log.err").Length -gt 0) {
    Write-Host '--- child stderr ---'
    Get-Content "$log.err" | ForEach-Object { "  $_" }
}
$txt = Read-Shared $res
if ($txt -notmatch '=== probe done ===') {
    Write-Host 'PROBE DID NOT FINISH -- the last START line in the result file is where it stopped'
}
# leave no owned Excel behind under any exit path
$owned = (Read-Shared $pidFile).Trim()
foreach ($id in ($owned -split ',' | Where-Object { $_ })) {
    if (Get-Process -Id ([int]$id) -ErrorAction SilentlyContinue) { Stop-Process -Id ([int]$id) -Force }
}
Write-Host "result file: $res"
