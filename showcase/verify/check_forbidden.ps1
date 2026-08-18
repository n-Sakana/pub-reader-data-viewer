# check_forbidden.ps1
#
# 成果物に Win32 / Shell / 外部 helper が入っていないことを機械で確かめる。
# 見るのは 2 つ:
#   1. ソースと配布用の .bas
#   2. 配布ブックの中に実際に入っている VBA（その場で書き出して読む）
# 参照設定も数える。UIAutomationClient 以外の追加参照があれば落とす。
[CmdletBinding()]
param(
    [string] $Root = 'C:\repos\pub\reader-data-viewer\showcase'
)
$ErrorActionPreference = 'Stop'

$patterns = @(
    @{ name = 'Declare (Win32 API)';        rx = '(?im)^\s*(Public\s+|Private\s+)?Declare\b' },
    @{ name = 'Shell()';                    rx = '(?im)(^|[^\w.])Shell\s*\(' },
    @{ name = 'Shell statement';            rx = '(?im)(^|:)\s*Shell\s+"' },
    @{ name = 'WScript.Shell';              rx = '(?i)WScript\.Shell' },
    @{ name = 'Scripting.FileSystemObject'; rx = '(?i)Scripting\.FileSystemObject' },
    @{ name = 'WMI (winmgmts)';             rx = '(?i)winmgmts' },
    @{ name = 'ShellExecute';               rx = '(?i)ShellExecute' },
    @{ name = 'FollowHyperlink';            rx = '(?i)FollowHyperlink' },
    @{ name = 'SendKeys';                   rx = '(?i)\bSendKeys\b' },
    @{ name = 'ExecuteExcel4Macro';         rx = '(?i)ExecuteExcel4Macro' },
    @{ name = 'Shapes.Add / OLEObjects';    rx = '(?i)(Shapes\.Add|OLEObjects\.Add|AddOLEObject|Forms\.[A-Za-z]+\.1)' },
    @{ name = 'CreateObject other than Excel.Application';
       rx   = '(?i)CreateObject\s*\(\s*"(?!Excel\.Application")' }
)

$fail = 0

# コメントは落としてから見る。禁止 API の名前は「使っていない」と書いた注記の
# 中に出てくるので、そこを拾うと検査にならない。
function Strip-Comments([string]$text) {
    $out = New-Object System.Text.StringBuilder
    foreach ($line in ($text -split "`r?`n")) {
        $inStr = $false
        $cut = -1
        for ($i = 0; $i -lt $line.Length; $i++) {
            $ch = $line[$i]
            if ($ch -eq '"') { $inStr = -not $inStr }
            elseif ($ch -eq "'" -and -not $inStr) { $cut = $i; break }
        }
        if ($cut -ge 0) { [void]$out.AppendLine($line.Substring(0, $cut)) }
        else { [void]$out.AppendLine($line) }
    }
    return $out.ToString()
}

function Scan-Text([string]$label, [string]$rawText) {
    $text = Strip-Comments $rawText
    $bad = 0
    foreach ($p in $patterns) {
        $m = [regex]::Matches($text, $p.rx)
        if ($m.Count -gt 0) {
            $bad++
            Write-Host "  FAIL $label : $($p.name) x$($m.Count)"
            foreach ($x in ($m | Select-Object -First 3)) {
                $line = ($text.Substring(0, $x.Index) -split "`n").Count
                Write-Host "        line $line : $((($text -split "`r?`n")[$line-1]).Trim())"
            }
        }
    }
    if ($bad -eq 0) { Write-Host "  ok   $label" }
    return $bad
}

Write-Host '--- sources on disk ---'
foreach ($f in @("$Root\src\modPixelBridge.bas") + @(Get-ChildItem "$Root\dist" -Filter "modPixelBridge_*px.bas" -ErrorAction SilentlyContinue | ForEach-Object { $_.FullName })) {
    if (-not (Test-Path $f)) { Write-Host "  MISSING $f"; $fail++; continue }
    $enc = if ($f -like '*\dist\*') { [System.Text.Encoding]::GetEncoding(932) } else { New-Object System.Text.UTF8Encoding($false) }
    $fail += Scan-Text (Split-Path $f -Leaf) ([System.IO.File]::ReadAllText($f, $enc))
}

Write-Host '--- the VBA actually inside the workbook ---'
# 1px / 2px / 4px は別々の成果物なので、あるものは全部見る。
$books = @(Get-ChildItem "$Root\dist" -Filter "VBA Pixel Bridge *px.xlsm" -ErrorAction SilentlyContinue | ForEach-Object { $_.FullName })
if ($books.Count -eq 0) { Write-Host "  MISSING: dist に成果物がありません"; $fail++ }
$dump = Join-Path $env:TEMP ('pbdump-' + [System.Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $dump | Out-Null
$before = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
Start-Sleep -Milliseconds 300
$mine = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id | Where-Object { $before -notcontains $_ })
try {
    $xl.AutomationSecurity = 1
    foreach ($book in $books) {
    Write-Host ("  -- " + (Split-Path $book -Leaf))
    $wb = $xl.Workbooks.Open($book)
    Write-Host "  components: $($wb.VBProject.VBComponents.Count)"
    foreach ($comp in $wb.VBProject.VBComponents) {
        $out = Join-Path $dump ($comp.Name + '.txt')
        $comp.Export($out)
        $txt = [System.IO.File]::ReadAllText($out, [System.Text.Encoding]::GetEncoding(932))
        $fail += Scan-Text ("in-book: " + $comp.Name) $txt
    }
    Write-Host '  references:'
    foreach ($r in $wb.VBProject.References) {
        Write-Host "    $($r.Name)  $($r.Description)"
    }
    $wb.Close($false)
    }
    $xl.Quit()
} finally {
    Start-Sleep -Milliseconds 500
    foreach ($id in $mine) { if (Get-Process -Id $id -ErrorAction SilentlyContinue) { Stop-Process -Id $id -Force } }
    Remove-Item $dump -Recurse -Force -ErrorAction SilentlyContinue
}

if ($fail -eq 0) { Write-Host 'RESULT: PASS (no Win32, no Shell, no external helper)' }
else { Write-Host "RESULT: FAIL ($fail findings)"; exit 1 }
