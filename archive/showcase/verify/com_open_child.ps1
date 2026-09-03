# com_open_child.ps1 -- VERIFICATION ONLY. Opens the SAME distributable in an
# automation Excel, which loads no add-ins, and times the open and the screen
# build. Compared against the shell open, this says whether the cost belongs to
# the workbook or to the add-ins the normal start loads.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string] $Book,
    [Parameter(Mandatory=$true)][string] $Res,
    [Parameter(Mandatory=$true)][string] $PidFile,
    [switch] $Visible
)
$ErrorActionPreference = 'Stop'
function Note([string]$s) { Add-Content -Path $Res -Value $s -Encoding utf8; Write-Host $s }

$before = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
$xl = New-Object -ComObject Excel.Application
$xl.Visible = [bool]$Visible
$xl.DisplayAlerts = $false
Start-Sleep -Milliseconds 300
$mine = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id |
          Where-Object { $before -notcontains $_ })
Set-Content -Path $PidFile -Value ($mine -join ',') -Encoding ascii
Note "owned pid: $($mine -join ',')  addins=$($xl.AddIns.Count) com=$($xl.COMAddIns.Count)"

$sw = [System.Diagnostics.Stopwatch]::StartNew()
$xl.AutomationSecurity = 1
$wb = $xl.Workbooks.Open($Book)
Note ("open_ms {0}" -f $sw.ElapsedMilliseconds)

$sw.Restart()
$ping = $xl.Run('PbPing')
Note ("ping_ms {0}  [{1}]" -f $sw.ElapsedMilliseconds, $ping)

$sw.Restart()
$xl.Run('PbShow')
Note ("pbshow_ms {0}" -f $sw.ElapsedMilliseconds)

$sw.Restart()
$txt = $xl.Run('PbGet', 'pb_elapsed')
Note ("read_back [{0}] in {1} ms" -f $txt, $sw.ElapsedMilliseconds)

Start-Sleep -Seconds 3
Note "DONE"
try { $xl.Run('PbShutdown') } catch { Note "shutdown err: $($_.Exception.Message)" }
try { $wb.Close($false) } catch { }
try { $xl.Quit() } catch { }
Start-Sleep -Milliseconds 800
foreach ($id in $mine) { if (Get-Process -Id $id -ErrorAction SilentlyContinue) { Stop-Process -Id $id -Force } }
