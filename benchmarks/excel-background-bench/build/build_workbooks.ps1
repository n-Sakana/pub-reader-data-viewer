# ============================================================================
# build_workbooks.ps1 -- assemble ZipBench.xlsm and engine\ZipEngine.xlsm from
#                        the plain-text modules in vba\.
#
# The .bas files are the source of truth; the two .xlsm files are build output.
# That keeps the whole thing reviewable as text and reproducible from scratch.
#
# Three things this script has to get right:
#
#  1. Encoding. VBE imports a .bas in the system ANSI code page (CP932 here),
#     but the repo keeps the sources in UTF-8 so they read correctly anywhere.
#     Every module is transcoded to CP932 into a temp folder first, and the
#     script refuses to continue if any character would not survive.
#
#  2. Trust. Importing code requires "Trust access to the VBA project object
#     model" (HKCU\...\Excel\Security\AccessVBOM = 1). The script checks and
#     says so plainly instead of failing with a COM error nobody can read.
#
#  3. Visibility. Excel is driven invisibly, so a dialog nobody can see would
#     hang the build forever. Every step is timestamped into build_workbooks.log
#     so the last line always names where it stopped.
#
# It only ever touches the Excel instance it created itself.
#
#   .\build_workbooks.ps1
# ============================================================================
[CmdletBinding()]
param([string] $Root = "")
$ErrorActionPreference = 'Stop'
if ([string]::IsNullOrEmpty($Root)) { $Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path) }

$xlWBATWorksheet = -4167
$xlOpenXMLWorkbookMacroEnabled = 52

$buildLog = Join-Path $Root 'build\build_workbooks.log'
Set-Content -LiteralPath $buildLog -Value '' -Encoding utf8 -Force
function Step([string] $m) {
  $line = "{0}  {1}" -f (Get-Date -Format 'HH:mm:ss.fff'), $m
  Add-Content -LiteralPath $buildLog -Value $line -Encoding utf8
  Write-Output $line
}

# --- preflight -------------------------------------------------------------
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
(この設定はこのスクリプトがコードを取り込むためだけに要ります。
 出来上がった ZipBench.xlsm を動かすのには要りません。)
"@
}

# modZipUi holds every macro the buttons call. It must stay free of user-defined
# types: Application.Run cannot reach ANY procedure in a module that also exposes
# a public UDT signature, and ZbResult is passed all through the other modules.
$benchModules  = @('modZipRule', 'modZipBench', 'modZipVba', 'modZipUia',
                   'modZipWorker', 'modZipExcel', 'modZipXll', 'modZipEmit', 'modZipUi')
$engineModules = @('modZipRule', 'modZipEngine')

# --- regenerate the emitter first ------------------------------------------
# modZipEmit.bas is the copy of ZipWorker.cs / run_worker.bat / build_worker.ps1
# that Excel writes out at run time for the "emitted" worker. If it is not
# regenerated here, the workbook can quietly ship an older worker than the one
# in worker\, and the two variants stop being the same program. Generating it
# on every build makes that impossible.
& (Join-Path $Root 'build\gen_emitter.ps1') -Root $Root | Out-Null
Write-Output 'emitter regenerated from worker\ sources'

# --- name-collision preflight ----------------------------------------------
# VBA is case-insensitive, so `Public Const ZB_SHEET` and `Public Function
# ZB_Sheet()` are the SAME name: "ambiguous name detected". That is a project-
# wide compile error, and its only symptom through automation is every
# Application.Run failing with "macros may be disabled" -- which points nowhere
# near the real cause. Catch it here, in text, where it is obvious.
$declRe = '^\s*(?:Public\s+|Private\s+|Global\s+)?(?:Static\s+)?(?:Declare\s+(?:PtrSafe\s+)?)?(?:Sub|Function|Property\s+(?:Get|Let|Set)|Const|Type|Enum)\s+([A-Za-z_][A-Za-z0-9_]*)'
$varRe  = '^\s*(?:Public|Private|Global|Dim)\s+([A-Za-z_][A-Za-z0-9_]*)\s+As\s'
$collisions = @()
foreach ($m in ($benchModules + $engineModules | Sort-Object -Unique)) {
  $seen = @{}
  foreach ($line in [IO.File]::ReadAllLines((Join-Path $Root "vba\$m.bas"))) {
    if ($line -match "^\s*'") { continue }
    $name = $null
    if ($line -match $declRe) { $name = $Matches[1] }
    elseif ($line -match $varRe) { $name = $Matches[1] }
    if (-not $name) { continue }
    $k = $name.ToLowerInvariant()
    if ($seen.ContainsKey($k) -and $seen[$k] -ne $name) {
      $collisions += "$m : '$($seen[$k])' と '$name' は VBA では同じ名前 (大文字小文字を区別しない)"
    }
    $seen[$k] = $name
  }
}
if ($collisions.Count) {
  throw ("VBA の名前衝突があります。直してから再実行してください:`n  " + ($collisions -join "`n  "))
}
Write-Output 'name-collision preflight: ok'

# --- one-line procedure guard ----------------------------------------------
# "Public Function F() As T: F = x: End Function" compiles in some places and
# not others; when it fails, VBA says only "only comments may follow End Sub"
# and the whole project stops compiling, which surfaces through automation as
# a hidden dialog. Ban the form outright.
$oneLiners = @()
foreach ($m in ($benchModules + $engineModules | Sort-Object -Unique)) {
  $ln = 0
  foreach ($line in [IO.File]::ReadAllLines((Join-Path $Root "vba\$m.bas"))) {
    $ln++
    if ($line -match "^\s*'") { continue }
    if ($line -match '(?i)^\s*(Public|Private|Friend)?\s*(Sub|Function|Property)\b.*:\s*End\s+(Sub|Function|Property)\s*$') {
      $oneLiners += "$m line ${ln}: $($line.Trim())"
    }
  }
}
if ($oneLiners.Count) {
  throw ("1 行に詰めた手続き定義があります。複数行に展開してください:`n  " + ($oneLiners -join "`n  "))
}
Write-Output 'one-line procedure guard: ok'
# --- declarations-after-procedure guard ------------------------------------
# VBA requires every module-level declaration (Dim/Const/Type/Declare/Public)
# to sit above the first procedure. Putting one after an End Sub gives the same
# opaque "only comments may follow End Sub" error, project-wide, with no line
# number reachable from automation. Catch it here.
$misplaced = @()
foreach ($m in ($benchModules + $engineModules | Sort-Object -Unique)) {
  $ln = 0; $seenProc = $false
  foreach ($line in [IO.File]::ReadAllLines((Join-Path $Root "vba\$m.bas"))) {
    $ln++
    if ($line -match "^\s*'") { continue }
    if ($line -match '(?i)^\s*(Public\s+|Private\s+|Friend\s+)?(Sub|Function|Property)\s') { $seenProc = $true; continue }
    if (-not $seenProc) { continue }
    if ($line -match '(?i)^\s*(Public|Private|Global)\s+(Const|Type|Declare|WithEvents)\s' -or
        $line -match '(?i)^\s*(Public|Private|Global)\s+[A-Za-z_][A-Za-z0-9_]*\s+As\s' -or
        $line -match '(?i)^\s*Declare\s') {
      $misplaced += "$m line ${ln}: $($line.Trim())"
    }
  }
}
if ($misplaced.Count) {
  throw ("モジュールレベルの宣言が手続きより後ろにあります。宣言部へ移してください:`n  " + ($misplaced -join "`n  "))
}
Write-Output 'declaration-order guard: ok'
# --- transcode UTF-8 -> CP932 ---------------------------------------------
$stage = Join-Path $env:TEMP ("zipbench_bas_" + [Guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $stage | Out-Null
$cp932 = [Text.Encoding]::GetEncoding(932)
foreach ($m in ($benchModules + $engineModules | Sort-Object -Unique)) {
  $src = Join-Path $Root "vba\$m.bas"
  if (-not (Test-Path -LiteralPath $src)) { throw "missing module: $src" }
  $t = [IO.File]::ReadAllText($src)
  if ($t -ne $cp932.GetString($cp932.GetBytes($t))) {
    throw "$src has characters that cannot survive CP932; VBE would import them as '?'"
  }
  [IO.File]::WriteAllText((Join-Path $stage "$m.bas"), $t, $cp932)
}
Step "staged modules as CP932 in $stage"

# --- Excel -----------------------------------------------------------------
$before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AskToUpdateLinks = $false
$xl.AlertBeforeOverwriting = $false
$after = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$mine = @($after | Where-Object { $before -notcontains $_ })
Step ("started Excel (new pid: {0})" -f ($(if ($mine.Count) { $mine -join ',' } else { 'reused' })))
if ($mine.Count) { Set-Content -LiteralPath "$buildLog.pid" -Value ($mine -join ',') }

function Import-Modules($wb, $names) {
  foreach ($n in $names) {
    [void]$wb.VBProject.VBComponents.Import((Join-Path $stage "$n.bas"))
    Step "    imported $n"
  }
}

try {
  # ---------- ZipEngine.xlsm (方式7 が別プロセスで開くエンジン) ----------
  $engineDir = Join-Path $Root 'engine'
  if (-not (Test-Path $engineDir)) { New-Item -ItemType Directory $engineDir | Out-Null }
  $enginePath = Join-Path $engineDir 'ZipEngine.xlsm'
  if (Test-Path $enginePath) { Remove-Item -LiteralPath $enginePath -Force }

  Step 'engine: Workbooks.Add'
  $wbE = $xl.Workbooks.Add($xlWBATWorksheet)
  Step 'engine: SaveAs'
  $wbE.SaveAs($enginePath, $xlOpenXMLWorkbookMacroEnabled)
  Step 'engine: import'
  Import-Modules $wbE $engineModules
  $wsE = $wbE.Worksheets.Item(1)
  $wsE.Name = 'ENGINE'
  $wsE.Range('A1').Value2 = 'ZipEngine -- ZipBench 方式7 が不可視の別 Excel プロセスで開くエンジン'
  $wsE.Range('A2').Value2 = 'このブックは単体では何もしない。ZipBench.xlsm が Application.Run で呼ぶ。'
  $wsE.Range('A3').Value2 = '変換規則は ZipBench.xlsm と同じ modZipRule.bas をそのまま取り込んである。'
  $wsE.Columns('A').ColumnWidth = 90
  Step 'engine: Save'
  $wbE.Save()
  $wbE.Close($false)
  Step "engine: built $enginePath"

  # ---------- ZipBench.xlsm (画面) ----------
  $benchPath = Join-Path $Root 'ZipBench.xlsm'
  if (Test-Path $benchPath) { Remove-Item -LiteralPath $benchPath -Force }

  Step 'bench: Workbooks.Add'
  $wbB = $xl.Workbooks.Add($xlWBATWorksheet)
  # Save first: Application.Run resolves a macro through the workbook's name, and
  # an unsaved book called "Sheet1" collides with its own sheet name.
  Step 'bench: SaveAs'
  $wbB.SaveAs($benchPath, $xlOpenXMLWorkbookMacroEnabled)

  Step 'bench: import'
  Import-Modules $wbB $benchModules

  # UIA のクライアント側は IUnknown 派生で IDispatch を持たない。
  # 遅延バインドでは呼び出せないので、参照設定は必須。
  #
  # 追加は「モジュールを取り込んだ後」でなければならない。先に参照を足すと、
  # そのあとに Import したモジュールの手続きが Application.Run から見えなくなり、
  # 「マクロを実行できません」になる (実測: Excel 16.0 / Win11 26200)。
  Step 'bench: AddFromGuid UIAutomationClient'
  [void]$wbB.VBProject.References.AddFromGuid('{944DE083-8FB8-45CF-BCB7-C477ACB2F897}', 1, 0)

  # No Application.Run during the build. xltoolrack's Build-Addin.ps1 never runs a
  # macro either: it writes ThisWorkbook code and lets the workbook initialise
  # itself on Workbook_Open. Running macros mid-build is where this script kept
  # stalling on invisible dialogs, and it buys nothing.
  Step 'bench: sheets + ThisWorkbook code'
  $ws = $wbB.Worksheets.Item(1)
  $ws.Name = 'BENCH'
  $wsLog = $wbB.Worksheets.Add([Type]::Missing, $ws)
  $wsLog.Name = 'LOG'
  $wsSig = $wbB.Worksheets.Add([Type]::Missing, $wsLog)
  $wsSig.Name = 'SIGNAL'
  $wsRuns = $wbB.Worksheets.Add([Type]::Missing, $wsSig)
  $wsRuns.Name = 'RUNS'
  $wsSig.Range('A3').Value2 = 'C#ワーカーはここへ完了を書く。その Worksheet_Change が完了検知そのもの。'
  $wsSig.Range('A4').Value2 = '電文は  DONE|件数|ハッシュ|変換ms|書込ms|接続ms'
  $wsSig.Columns('A').ColumnWidth = 80

  # The completion event. This is the whole detection mechanism for the C#
  # worker methods: no polling, no file, one cell change.
  # The VB component is named after the sheet's CodeName ("Sheet3"), not after
  # the tab name, so ask the sheet which component belongs to it.
  $sigMod = $wbB.VBProject.VBComponents.Item($wsSig.CodeName).CodeModule
  if ($sigMod.CountOfLines -gt 0) { $sigMod.DeleteLines(1, $sigMod.CountOfLines) }
  $sigMod.AddFromString(@'
Option Explicit

' Every method updates this one cell once, at the very end. That single change
' is how Excel learns the run is finished -- the C# worker writes it over COM,
' the in-process methods write it themselves. The end time is taken inside this
' handler, so however coarsely the caller happens to be pumping messages, it
' cannot show up in the measurement.
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    If Intersect(Target, Me.Range("A1")) Is Nothing Then Exit Sub
    modZipBench.ZbSignalFired CStr(Me.Range("A1").Value)
End Sub
'@)

  $tw = $wbB.VBProject.VBComponents.Item('ThisWorkbook').CodeModule
  if ($tw.CountOfLines -gt 0) { $tw.DeleteLines(1, $tw.CountOfLines) }
  $tw.AddFromString(@'
Option Explicit

' The builder never runs a macro. The screen is painted the first time this
' workbook is opened normally, which is also the only way a user ever sees it.
Private Sub Workbook_Open()
    On Error Resume Next
    modZipUi.ZB_InitSheetIfEmpty
End Sub

' Close down anything this workbook started: its own worker process, its own
' hidden Excel, its own scratch workbook. Nothing else is touched.
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    modZipUi.ZB_Cleanup
    Me.Saved = True
End Sub
'@)

  Step 'bench: buttons'
  $anchor = $ws.Range('AD3')
  $left = [double]$anchor.Left
  $y = [double]$anchor.Top
  $w = 168.0
  $h = 22.0
  $gap = 3.0

  $buttons = @(
    @{ c = 'KEN_ALL を取得';              m = 'ZB_DownloadKenAll' },
    @{ c = '準備 (辞書・入力・正解)';      m = 'ZB_Prepare' },
    @{ c = '';                             m = '' },
    @{ c = '1  低速VBA (1セルずつ)';      m = 'ZB_Run1' },
    @{ c = '2  高速VBA (配列一括)';       m = 'ZB_Run2' },
    @{ c = '3  BAT手動起動';              m = 'ZB_Run3' },
    @{ c = '4  WScript.Shell.Run';        m = 'ZB_Run4' },
    @{ c = '5  VBA Shell';                m = 'ZB_Run5' },
    @{ c = '6  Task Scheduler COM';       m = 'ZB_Run6' },
    @{ c = '7  不可視の別Excel';          m = 'ZB_Run7' },
    @{ c = '8  ShellExecute';             m = 'ZB_Run8' },
    @{ c = '9  候補2 TSV+QueryTable';     m = 'ZB_Run9' },
    @{ c = '10 候補3 COM Value2';         m = 'ZB_Run10' },
    @{ c = '13 ADO CopyFromRecordset';    m = 'ZB_Run13' },
    @{ c = '14 DAO CopyFromRecordset';    m = 'ZB_Run14' },
    @{ c = '15 Workbooks.OpenText';       m = 'ZB_Run15' },
    @{ c = '16 Excel-DNA XLL';            m = 'ZB_Run16' },
    @{ c = '';                             m = '' },
    @{ c = '全方式を順に実行';             m = 'ZB_RunAll' },
    @{ c = '';                             m = '' },
    @{ c = 'ファイル作成だけ (prebuilt)';  m = 'ZB_EmitOnlyPrebuilt' },
    @{ c = 'ファイル作成だけ (emitted)';   m = 'ZB_EmitOnlyEmitted' },
    @{ c = '';                             m = '' },
    @{ c = '取消';                         m = 'ZB_Cancel' },
    @{ c = '後始末';                       m = 'ZB_Cleanup' },
    @{ c = '画面を作り直す';               m = 'ZB_InitSheet' }
  )

  Step 'bench: adding buttons'
  foreach ($b in $buttons) {
    if ([string]::IsNullOrEmpty($b.m)) { $y += $h * 0.5; continue }
    $btn = $ws.Buttons().Add($left, $y, $w, $h)
    $btn.Caption = $b.c
    $btn.OnAction = $b.m
    $y += $h + $gap
  }

  $ws.Range('AD1').Value2 = '操作'
  $ws.Range('AD1').Font.Bold = $true

  Step 'bench: final Save'
  $wbB.Save()
  $wbB.Close($false)
  Step "bench: built $benchPath"
}
catch {
  Step ("FAILED: " + $_.Exception.Message)
  Step ("        at " + $_.InvocationInfo.PositionMessage.Trim())
  throw
}
finally {
  Step 'quitting Excel'
  try { $xl.Quit() } catch { }
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  Remove-Item -LiteralPath $stage -Recurse -Force -ErrorAction SilentlyContinue
}
Step 'done'
