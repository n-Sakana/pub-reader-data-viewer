param([string]$Book, [string]$Res, [switch]$Compile)
$ErrorActionPreference = 'Continue'
$out = New-Object System.Collections.ArrayList
function Note($m) { [void]$out.Add([string]$m); [System.IO.File]::WriteAllLines($Res, $out) }

$dir  = [System.IO.Path]::Combine($env:TEMP, 'pixelbridge')
$prog = [System.IO.Path]::Combine($dir, 'progress.json')
$cmd  = [System.IO.Path]::Combine($dir, 'command.json')
$fe   = [System.IO.Path]::Combine($dir, 'fe.json')
if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
Remove-Item -LiteralPath $prog, $cmd -Force -ErrorAction SilentlyContinue
[System.IO.File]::WriteAllText($fe, '{"state":"probe"}')

function Peek {
    try {
        $fs = [System.IO.File]::Open($prog, 'Open', 'Read', 'ReadWrite')
        $sr = New-Object System.IO.StreamReader($fs)
        $t = $sr.ReadToEnd(); $sr.Close(); $fs.Close(); return $t.Trim()
    } catch { return '' }
}
function Watch($seconds, $tag) {
    $last = ''
    $dl = (Get-Date).AddSeconds($seconds)
    while ((Get-Date) -lt $dl) {
        $t = Peek
        if ($t -and $t -ne $last) {
            $last = $t
            Note ("  [$tag] " + $(if ($t.Length -gt 110) { $t.Substring(0,110) + '...' } else { $t }))
        }
        Start-Sleep -Milliseconds 250
    }
    return $last
}

$before = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$xl = $null
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.ScreenUpdating = $false
    $xl.EnableEvents = $false
    $xl.AutomationSecurity = 1
    $wb = $xl.Workbooks.Open($Book, 0, $true)
    $q = "'" + $wb.Name + "'!"
    if ($Compile) {
        # inspect_be.ps1 がやっていた「VBE で VBAProject をコンパイル」を再現する
        $vbe = $xl.VBE
        $vbe.ActiveVBProject = $wb.VBProject
        $ctl = $vbe.CommandBars.FindControl(1, 578)
        if ($null -eq $ctl) { Note 'compile: NOCONTROL' }
        else { $ctl.Execute(); Note ('compile: executed, enabled=' + $ctl.Enabled) }
    }
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    Note ('PbBeBootstrap -> ' + $xl.Run($q + 'PbBeBootstrap') + ' in ' + $sw.ElapsedMilliseconds + ' ms')
    $mine = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object { $_.Id })
    Note ('BE pid(s) -> ' + ($mine -join ','))
    $xl = $null   # FE と同じく参照を落とす。UserControl=True が効いているかも見る

    Note 'watching the resident loop for 5s (no COM held):'
    $idle = Watch 5 'idle'
    Note ('idle seen -> ' + [bool]$idle)

    Note 'sending pi n=134 cols=67:'
    [System.IO.File]::WriteAllText($cmd, '{"seq":1,"cmd":"pi","n":134,"cols":67}')
    $last = Watch 60 'pi'
    Note ('last -> ' + $(if ($last.Length -gt 160) { $last.Substring(0,160) + '...' } else { $last }))

    Note 'sending quit:'
    [System.IO.File]::WriteAllText($cmd, '{"seq":2,"cmd":"quit"}')
    Watch 10 'quit' | Out-Null
    Start-Sleep -Seconds 2
    $still = @(Get-Process -Id $mine -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    Note ('BE process still alive -> ' + ($still -join ',') + ($(if ($still.Count -eq 0) { ' (exited cleanly)' } else { ' (LEAK)' })))
} catch {
    Note ('ERROR ' + $_.Exception.Message)
}
Remove-Item -LiteralPath $fe -Force -ErrorAction SilentlyContinue
Get-Process -Name EXCEL -ErrorAction SilentlyContinue |
    Where-Object { $before -notcontains $_.Id } |
    ForEach-Object { Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue }
Note 'DONE'
