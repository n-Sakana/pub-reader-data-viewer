# ============================================================================
# build_workbooks.ps1 -- assemble the two workbooks from src\vba.
#
#   dist\ReaderDataViewer-VBA.xlsm      method 1: everything inside Excel
#   dist\ReaderDataViewer-Hybrid.xlsm   method 3: the display half only
#
# The .bas files are the source of truth; the .xlsm files are build output.
# Both books get the SAME sheet layout, painted by the same function below, so
# the three methods really are showing the same thing in the same places.
#
# Four things this has to get right, all of them learned the hard way in
# benchmarks\excel-background-bench:
#
#  1. Encoding. VBE imports a .bas in the system ANSI code page (CP932 here)
#     while the repo keeps sources in UTF-8. Every module is transcoded first
#     and the build stops if a character would not survive the round trip.
#
#  2. Trust. Importing code needs "Trust access to the VBA project object
#     model". Checked up front, in words, instead of failing with a COM error.
#
#  3. Reference order. UIAutomationClient must be added AFTER the modules are
#     imported. Adding it first makes every Application.Run fail with "macros
#     may be disabled", which points nowhere near the real cause.
#
#  4. Compile traps. VBA is case-insensitive, will not take a one-line
#     procedure body, and wants module-level declarations above the first
#     procedure. Each of those produces the same unhelpful dialog, so all three
#     are checked as text before Excel ever opens.
#
# Plus one specific to this app: the polynomial hash is inlined into three hot
# loops because a VBA call costs more than the hash. Those copies are compared
# against the canonical RdvHashBytes here, so they cannot drift apart.
#
# It only ever touches the Excel instance it created itself.
# ============================================================================
[CmdletBinding()]
param([string] $Root = "")
$ErrorActionPreference = 'Stop'
if ([string]::IsNullOrEmpty($Root)) { $Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path) }

$xlWBATWorksheet = -4167
$xlOpenXMLWorkbookMacroEnabled = 52

$buildLog = Join-Path $Root 'work\build_workbooks.log'
$workDir = Split-Path -Parent $buildLog
if (-not (Test-Path -LiteralPath $workDir)) { New-Item -ItemType Directory -Path $workDir | Out-Null }
Set-Content -LiteralPath $buildLog -Value '' -Encoding utf8 -Force
function Step([string] $m) {
  $line = "{0}  {1}" -f (Get-Date -Format 'HH:mm:ss.fff'), $m
  Add-Content -LiteralPath $buildLog -Value $line -Encoding utf8
  Write-Output $line
}

$modules = @('modRdvSpec', 'modRdvUia', 'modRdvEngine', 'modRdvApp')

# --- preflight: trust -------------------------------------------------------
$vbom = 0
foreach ($v in @('16.0', '15.0', '14.0')) {
  $p = "HKCU:\Software\Microsoft\Office\$v\Excel\Security"
  if (Test-Path $p) { $x = (Get-ItemProperty $p).AccessVBOM; if ($x) { $vbom = $x } }
}
if ($vbom -ne 1) {
  throw @"
Excel の「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」が無効です。
  Excel > ファイル > オプション > トラスト センター > トラスト センターの設定
        > マクロの設定 > VBA プロジェクト オブジェクト モデルへのアクセスを信頼する
にチェックを入れてから、もう一度このスクリプトを実行してください。
(この設定はビルドがコードを取り込むためだけに要ります。
 出来上がった .xlsm を動かすのには要りません。)
"@
}

# --- preflight: name collisions --------------------------------------------
$declRe = '^\s*(?:Public\s+|Private\s+|Global\s+)?(?:Static\s+)?(?:Declare\s+(?:PtrSafe\s+)?)?(?:Sub|Function|Property\s+(?:Get|Let|Set)|Const|Type|Enum)\s+([A-Za-z_][A-Za-z0-9_]*)'
# Every declared name on the line, arrays and comma lists included. The first
# version of this check only looked at the first "<name> As" on a line, and so
# missed "Private m_bC() As Byte" colliding with "Private m_bc() As Long" --
# the same name to VBA. That is a project-wide compile error whose only symptom
# through automation is an invisible dialog and "macros may be disabled".
$varRe = '(?<![A-Za-z0-9_])([A-Za-z_][A-Za-z0-9_]*)\s*(?:\([^)]*\))?\s+As\s'
$collisions = @()
foreach ($m in $modules) {
  $seen = @{}
  foreach ($line in [IO.File]::ReadAllLines((Join-Path $Root "src\vba\$m.bas"))) {
    if ($line -match "^\s*'") { continue }
    $names = @()
    if ($line -match $declRe) { $names += $Matches[1] }
    if ($line -match '(?i)^\s*(Public|Private|Global|Dim)\s' -and
        $line -notmatch '(?i)^\s*(Public|Private|Global)?\s*(Declare|Sub|Function|Property|Const|Type|Enum)\b') {
      foreach ($mm in [regex]::Matches($line, $varRe)) { $names += $mm.Groups[1].Value }
    }
    foreach ($name in $names) {
      $k = $name.ToLowerInvariant()
      if ($seen.ContainsKey($k) -and $seen[$k] -ne $name) {
        $collisions += "$m : '$($seen[$k])' と '$name' は VBA では同じ名前 (大文字小文字を区別しない)"
      }
      $seen[$k] = $name
    }
  }
}
if ($collisions.Count) { throw ("VBA の名前衝突:`n  " + ($collisions -join "`n  ")) }
Write-Output 'name-collision preflight: ok'

# --- preflight: one-line procedures + declaration order ----------------------
$oneLiners = @()
$misplaced = @()
foreach ($m in $modules) {
  $ln = 0; $seenProc = $false
  foreach ($line in [IO.File]::ReadAllLines((Join-Path $Root "src\vba\$m.bas"))) {
    $ln++
    if ($line -match "^\s*'") { continue }
    if ($line -match '(?i)^\s*(Public|Private|Friend)?\s*(Sub|Function|Property)\b.*:\s*End\s+(Sub|Function|Property)\s*$') {
      $oneLiners += "$m line ${ln}: $($line.Trim())"
    }
    if ($line -match '(?i)^\s*(Public\s+|Private\s+|Friend\s+)?(Sub|Function|Property)\s') { $seenProc = $true; continue }
    if (-not $seenProc) { continue }
    if ($line -match '(?i)^\s*(Public|Private|Global)\s+(Const|Type|Declare|WithEvents)\s' -or
        $line -match '(?i)^\s*(Public|Private|Global)\s+[A-Za-z_][A-Za-z0-9_]*\s+As\s' -or
        $line -match '(?i)^\s*Declare\s') {
      $misplaced += "$m line ${ln}: $($line.Trim())"
    }
  }
}
if ($oneLiners.Count) { throw ("1 行に詰めた手続き定義:`n  " + ($oneLiners -join "`n  ")) }
if ($misplaced.Count) { throw ("宣言が手続きより後ろ:`n  " + ($misplaced -join "`n  ")) }
Write-Output 'one-line / declaration-order preflight: ok'

# --- preflight: the inlined hash still matches the canonical one -------------
function Normalize-Hash([string[]] $lines) {
  $out = @()
  foreach ($l in $lines) {
    $t = $l.Trim()
    if ($t.Length -eq 0) { continue }
    if ($t.StartsWith("'")) { continue }
    # the copies read from differently named arrays; the arithmetic must match
    $t = [regex]::Replace($t, '(?<![A-Za-z0-9_])[A-Za-z_][A-Za-z0-9_]*\(p', 'ARR(p')
    $out += $t
  }
  return ($out -join '|')
}
$specText = [IO.File]::ReadAllLines((Join-Path $Root 'src\vba\modRdvSpec.bas'))
$canon = @()
$inFn = $false
foreach ($l in $specText) {
  if ($l -match '^\s*Public Function RdvHashBytes') { $inFn = $true; continue }
  if ($inFn) {
    if ($l -match '^\s*End Function') { break }
    if ($l -match '^\s*Dim ' -or $l -match '^\s*RdvHashBytes\s*=') { continue }
    $canon += $l
  }
}
$canonKey = Normalize-Hash $canon
if ($canonKey.Length -lt 50) { throw 'canonical RdvHashBytes body not found' }

$engText = [IO.File]::ReadAllLines((Join-Path $Root 'src\vba\modRdvEngine.bas'))
$copies = 0; $block = @(); $inBlock = $false
foreach ($l in $engText) {
  if ($l -match "RdvHashBytes, inlined") { $inBlock = $true; $block = @(); continue }
  if ($inBlock) {
    if ($l -match 'end inline') {
      $inBlock = $false
      $copies++
      if ((Normalize-Hash $block) -ne $canonKey) {
        throw "modRdvEngine の inline ハッシュ #$copies が modRdvSpec.RdvHashBytes と一致しません"
      }
      continue
    }
    $block += $l
  }
}
if ($copies -ne 3) { throw "inline ハッシュは 3 か所のはずが $copies か所でした" }
Write-Output "inline-hash preflight: ok ($copies copies identical to RdvHashBytes)"

# --- transcode UTF-8 -> CP932 ----------------------------------------------
$stage = Join-Path $env:TEMP ("rdv_bas_" + [Guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $stage | Out-Null
$cp932 = [Text.Encoding]::GetEncoding(932)
foreach ($m in $modules) {
  $src = Join-Path $Root "src\vba\$m.bas"
  $t = [IO.File]::ReadAllText($src)
  if ($t -ne $cp932.GetString($cp932.GetBytes($t))) {
    throw "$src has characters that cannot survive CP932; VBE would import them as '?'"
  }
  [IO.File]::WriteAllText((Join-Path $stage "$m.bas"), $t, $cp932)
}
Step "staged modules as CP932 in $stage"

# --- the shared screen ------------------------------------------------------
function Paint-View($ws, [string] $method, [bool] $withButtons) {
  $ws.Name = 'VIEW'
  $ws.Cells.Font.Name = 'Yu Gothic UI'
  $ws.Cells.Font.Size = 10

  $ws.Columns('A').ColumnWidth = 2
  $ws.Columns('B').ColumnWidth = 15
  $ws.Columns('C').ColumnWidth = 26
  $ws.Columns('D').ColumnWidth = 15
  $ws.Columns('E').ColumnWidth = 26
  $ws.Columns('F').ColumnWidth = 15
  $ws.Columns('G').ColumnWidth = 26
  $ws.Columns('H').ColumnWidth = 15
  $ws.Columns('I').ColumnWidth = 2
  $ws.Columns('J').ColumnWidth = 26
  $ws.Columns('K').ColumnWidth = 12
  $ws.Columns('L').ColumnWidth = 9

  $ws.Range('B1').Value2 = 'Reader Data Viewer'
  $ws.Range('B1').Font.Size = 18
  $ws.Range('B1').Font.Bold = $true
  $ws.Range('E1').Value2 = $method
  $ws.Range('E1').Font.Size = 12
  $ws.Range('E1').Font.Bold = $true
  $ws.Range('E1').Font.Color = 8210719

  $ws.Range('B2').Value2 = '状態'
  $ws.Range('B3').Value2 = 'メモ帳'
  $ws.Range('B4').Value2 = 'データ'
  $ws.Range('B2:B4').Font.Color = 8210719
  $ws.Range('C2').Font.Bold = $true
  $ws.Range('C2').Font.Size = 12

  $ws.Range('E2').Value2 = '手動実行の番号'
  $ws.Range('E2').Font.Color = 8210719
  $ws.Range('F2').NumberFormat = '@'
  $ws.Range('F2').Interior.Color = 15921906
  $ws.Range('F2').HorizontalAlignment = -4108

  $ws.Range('B6').Value2 = '現在の番号 (番号1)'
  $ws.Range('B6').Font.Color = 8210719
  $ws.Range('C6').Font.Size = 22
  $ws.Range('C6').Font.Bold = $true
  $ws.Range('C6').NumberFormat = '@'
  $ws.Range('E6').Font.Size = 12
  $ws.Range('E6').Font.Bold = $true
  $ws.Range('B6:H6').Interior.Color = 16250871

  $hdr = @('#', '表A 項目', '表A 値', '表B 項目', '表B 値', '表C 項目', '表C 値')
  for ($i = 0; $i -lt $hdr.Count; $i++) {
    $c = $ws.Cells.Item(8, 2 + $i)
    $c.Value2 = $hdr[$i]
    $c.Font.Bold = $true
    $c.Interior.Color = 14994616
  }
  $ws.Range('B9:H18').NumberFormat = '@'
  $ws.Range('B9:B18').HorizontalAlignment = -4152
  $ws.Range('B8:H18').Borders.LineStyle = 1
  $ws.Range('B8:H18').Borders.Color = 13421772

  $ws.Range('J8').Value2 = '工程'
  $ws.Range('K8').Value2 = 'ミリ秒'
  $ws.Range('L8').Value2 = '割合'
  $ws.Range('J8:L8').Font.Bold = $true
  $ws.Range('J8:L8').Interior.Color = 14994616
  $ws.Range('K9:K20').NumberFormat = '#,##0.0'
  $ws.Range('L9:L20').NumberFormat = '0.0"%"'
  $ws.Range('J8:L20').Borders.LineStyle = 1
  $ws.Range('J8:L20').Borders.Color = 13421772
  $ws.Range('J17').Font.Bold = $true
  $ws.Range('K17').Font.Bold = $true

  $h2 = @('#', '時刻', '番号1', '総時間ms', '検知ms', '照合')
  for ($i = 0; $i -lt $h2.Count; $i++) {
    $c = $ws.Cells.Item(21, 2 + $i)
    $c.Value2 = $h2[$i]
    $c.Font.Bold = $true
    $c.Interior.Color = 14994616
  }
  $ws.Range('D22:D41').NumberFormat = '@'
  $ws.Range('E22:F41').NumberFormat = '#,##0.0'
  $ws.Range('B21:G41').Borders.LineStyle = 1
  $ws.Range('B21:G41').Borders.Color = 13421772

  $ws.Range('B44').Value2 = 'エラー'
  $ws.Range('B45').Value2 = '停止フラグ'
  $ws.Range('B46').Value2 = 'ログ出力先'
  $ws.Range('B47').Value2 = 'データフォルダ'
  $ws.Range('B44:B47').Font.Color = 8210719
  $ws.Range('C44').Font.Color = 2701998
  $ws.Range('C45:C47').Font.Color = 8210719
  $ws.Range('B49').Value2 = 'メモ帳に 8 桁の番号を入力すると、3 本の CSV を読み直して結合し直し、この画面を更新します。'
  $ws.Range('B50').Value2 = '同じ番号を続けて読むときは、いったん欄を空にしてから読み取ってください。'
  $ws.Range('B49:B50').Font.Color = 8210719

  if ($withButtons) {
    $ws.Range('N1').Value2 = '操作'
    $ws.Range('N1').Font.Bold = $true
    $anchor = $ws.Range('N2')
    $left = [double]$anchor.Left
    $y = [double]$anchor.Top
    $buttons = @(
      @{ c = '監視開始'; m = 'RDV_StartMonitor' },
      @{ c = '監視停止'; m = 'RDV_StopMonitor' },
      @{ c = '手動実行 (F2 の番号)'; m = 'RDV_ManualRun' },
      @{ c = '画面クリア'; m = 'RDV_Reset' }
    )
    foreach ($b in $buttons) {
      $btn = $ws.Buttons().Add($left, $y, 170.0, 24.0)
      $btn.Caption = $b.c
      $btn.OnAction = $b.m
      $y += 28.0
    }
  }
  $ws.Range('A1').Select() | Out-Null
}

# --- Excel ------------------------------------------------------------------
$before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AskToUpdateLinks = $false
$xl.AlertBeforeOverwriting = $false
$after = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$mine = @($after | Where-Object { $before -notcontains $_ })
Step ("started Excel (new pid: {0})" -f $(if ($mine.Count) { $mine -join ',' } else { 'reused' }))

$dist = Join-Path $Root 'dist'
if (-not (Test-Path -LiteralPath $dist)) { New-Item -ItemType Directory -Path $dist | Out-Null }

try {
  # ---------- the display half of the hybrid (no code at all) ----------
  $hybPath = Join-Path $dist 'ReaderDataViewer-Hybrid.xlsm'
  if (Test-Path -LiteralPath $hybPath) { Remove-Item -LiteralPath $hybPath -Force }
  Step 'hybrid: Workbooks.Add'
  $wbH = $xl.Workbooks.Add($xlWBATWorksheet)
  $wbH.SaveAs($hybPath, $xlOpenXMLWorkbookMacroEnabled)
  Paint-View $wbH.Worksheets.Item(1) '併用方式 (C# + Excel)' $false
  $wsH = $wbH.Worksheets.Item(1)
  $wsH.Range('C2').Value2 = '未接続'
  $wsH.Range('B52').Value2 = 'この画面は表示専用です。読み取り・CSV 読込・結合・計測は ReaderDataViewer-Hybrid.cmd が行います。'
  $wsH.Range('B52').Font.Color = 8210719
  $wbH.Save()
  $wbH.Close($false)
  Step "hybrid: built $hybPath"

  # ---------- method 1: everything inside Excel ----------
  $vbaPath = Join-Path $dist 'ReaderDataViewer-VBA.xlsm'
  if (Test-Path -LiteralPath $vbaPath) { Remove-Item -LiteralPath $vbaPath -Force }
  Step 'vba: Workbooks.Add'
  $wbV = $xl.Workbooks.Add($xlWBATWorksheet)
  # save first: Application.Run resolves a macro through the workbook name, and
  # an unsaved book called "Sheet1" collides with its own sheet name
  $wbV.SaveAs($vbaPath, $xlOpenXMLWorkbookMacroEnabled)

  Step 'vba: import modules'
  foreach ($m in $modules) {
    [void]$wbV.VBProject.VBComponents.Import((Join-Path $stage "$m.bas"))
    Step "    imported $m"
  }

  # AFTER the imports, never before
  Step 'vba: AddFromGuid UIAutomationClient'
  [void]$wbV.VBProject.References.AddFromGuid('{944DE083-8FB8-45CF-BCB7-C477ACB2F897}', 1, 0)

  Paint-View $wbV.Worksheets.Item(1) 'VBA 単独方式' $true

  $tw = $wbV.VBProject.VBComponents.Item('ThisWorkbook').CodeModule
  if ($tw.CountOfLines -gt 0) { $tw.DeleteLines(1, $tw.CountOfLines) }
  $tw.AddFromString(@'
Option Explicit

' The builder never runs a macro. The screen is set up the first time the book
' is opened normally, which is the only way anyone ever sees it.
Private Sub Workbook_Open()
    On Error Resume Next
    modRdvApp.RDV_Reset
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    modRdvApp.RDV_StopMonitor
    Me.Saved = True
End Sub
'@)

  Step 'vba: final Save'
  $wbV.Save()
  $wbV.Close($false)
  Step "vba: built $vbaPath"
}
catch {
  Step ("FAILED: " + $_.Exception.Message)
  Step ("        at " + $_.InvocationInfo.PositionMessage.Trim())
  throw
}
finally {
  Step 'quitting Excel'
  try { $xl.Quit() } catch { }
  try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) } catch { }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  Start-Sleep -Milliseconds 500
  foreach ($p in $mine) {
    $q = Get-Process -Id $p -ErrorAction SilentlyContinue
    if ($q) { Step "closing my excel $p"; $q.Kill() }
  }
  Remove-Item -LiteralPath $stage -Recurse -Force -ErrorAction SilentlyContinue
}
Step 'done'
