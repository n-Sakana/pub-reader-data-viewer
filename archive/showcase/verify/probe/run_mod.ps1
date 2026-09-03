# run_mod.ps1 -- VERIFICATION ONLY. Watchdog for mod_child.ps1.
# Makes no COM call itself, kills only what the child recorded as its own, and
# never leaves anything on the operator's screen.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string] $Bas,
    [Parameter(Mandatory=$true)][string] $PingProc,
    [Parameter(Mandatory=$true)][string] $ArmProc,
    [Parameter(Mandatory=$true)][string] $DoneMark,
    [string] $Tag = 'mod',
    [string] $A2 = '',
    [switch] $Visible,
    [switch] $Uia,
    [int] $TimeoutSec = 150
)
$ErrorActionPreference = 'Stop'

$here    = Split-Path -Parent $MyInvocation.MyCommand.Path
$child   = Join-Path $here 'mod_child.ps1'
$res     = Join-Path $here "res-$Tag.txt"
$pidFile = Join-Path $here "owned-$Tag.pid"
$log     = Join-Path $here "log-$Tag.txt"
foreach ($f in @($res, $pidFile, $log, "$log.err")) { if (Test-Path $f) { Remove-Item $f -Force } }

function Read-Shared([string]$p) {
    if (-not (Test-Path $p)) { return '' }
    try {
        $fs = [System.IO.File]::Open($p, 'Open', 'Read', 'ReadWrite')
        $sr = New-Object System.IO.StreamReader($fs)
        $t  = $sr.ReadToEnd(); $sr.Close(); $fs.Close(); return $t
    } catch { return '' }
}
function Kill-Owned {
    foreach ($id in ((Read-Shared $pidFile).Trim() -split ',' | Where-Object { $_ })) {
        if (Get-Process -Id ([int]$id) -ErrorAction SilentlyContinue) {
            Write-Host "  killing owned excel pid $id"
            Stop-Process -Id ([int]$id) -Force
        }
    }
}

$childArgs = @('-ExecutionPolicy','Bypass','-NoProfile','-File',"`"$child`"",
               '-Bas',"`"$Bas`"",'-Res',"`"$res`"",'-PidFile',"`"$pidFile`"",
               '-PingProc',$PingProc,'-ArmProc',$ArmProc,'-DoneMark',"`"$DoneMark`"",
               '-WaitSec',[string]($TimeoutSec - 20))
if ($A2)      { $childArgs += @('-A2', "`"$A2`"") }
if ($Visible) { $childArgs += '-Visible' }
if ($Uia)     { $childArgs += '-Uia' }
$p = Start-Process -FilePath 'powershell.exe' -ArgumentList $childArgs -PassThru `
        -RedirectStandardOutput $log -RedirectStandardError "$log.err" -WindowStyle Hidden

$deadline = (Get-Date).AddSeconds($TimeoutSec)
while (-not $p.HasExited -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 400 }
if (-not $p.HasExited) {
    Write-Host 'DEADLINE: the child is stuck (a modal, or a call that never returns)'
    Kill-Owned
    Start-Sleep -Milliseconds 400
    if (-not $p.HasExited) { Stop-Process -Id $p.Id -Force }
}
Kill-Owned
Write-Host "--- child log ($Tag) ---"
if (Test-Path $log) { Get-Content $log | ForEach-Object { "  $_" } }
if ((Test-Path "$log.err") -and (Get-Item "$log.err").Length -gt 0) { Get-Content "$log.err" | ForEach-Object { "  ERR $_" } }
Write-Host "--- result ($Tag) ---"
$t = Read-Shared $res
if ($t) { $t -split "`r?`n" | Where-Object { $_ } | ForEach-Object { "  $_" } } else { '  (empty)' }
