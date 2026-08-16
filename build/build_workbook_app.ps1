# ============================================================================
# build_workbook_app.ps1 -- the practical VBA build, from src\app\vba.
#
#   dist\app-vba\ReaderDataViewer.xlsm          the small FE (UI + META only)
#   dist\app-vba\ReaderDataViewer-Ledger.xlsx   the BE-owned ledger workbook
#   dist\app-vba\ReaderDataViewer-Ledger.state  its sidecar mirror
#
# The FE holds NO ledger sheet and only five modules (Spec/Chan/Host/Ui/App +
# the events class); the heavy modules (Engine/Be/Uia) live only in the worker
# book embedded into META as base64. The initial ledger workbook + sidecar are
# seeded by running Rdv3BeBuildInitial inside the WORKER book -- the exact
# code path the product uses at apply time, not a builder re-implementation.
#
#   UI      the sheet-drawn screen: no UserForm, no ActiveX, no Forms controls,
#           no ListBox, no Shapes. Buttons are hyperlink cells (ScreenTip =
#           hover text) dispatched through Worksheet_FollowHyperlink.
#   META    hidden: overrides and the embedded worker book (base64).
#
# Same preflights as the frozen builders (CP932 round trip, CRLF staging, name
# collisions, declaration order), and it only ever touches the Excel instance
# it created itself.
# ============================================================================
[CmdletBinding()]
param([string] $Root = "")
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'excel_own.ps1')   # exact Excel ownership, never a pid diff
if ([string]::IsNullOrEmpty($Root)) { $Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path) }

$xlWBATWorksheet = -4167
$xlsm = 52

$buildLog = Join-Path $Root 'work\build_workbook_app.log'
$workDir = Split-Path -Parent $buildLog
if (-not (Test-Path -LiteralPath $workDir)) { New-Item -ItemType Directory -Path $workDir | Out-Null }
Set-Content -LiteralPath $buildLog -Value '' -Encoding utf8 -Force
function Step([string] $m) {
  $line = "{0}  {1}" -f (Get-Date -Format 'HH:mm:ss.fff'), $m
  Add-Content -LiteralPath $buildLog -Value $line -Encoding utf8
  Write-Output $line
}

# modRdv3Cfg is in BOTH books: the FE reads ReaderDataViewer.json for the paths
# and shows the settings, and the BE needs the key rule, the timings and the
# watch targets out of the same file.
$modules = @('modRdv3Spec', 'modRdv3Chan', 'modRdv3Cfg', 'modRdv3Host', 'modRdv3Ui', 'modRdv3Set', 'modRdv3App')
$workerModules = @('modRdv3Spec', 'modRdv3Chan', 'modRdv3Cfg', 'modRdv3Uia', 'modRdv3Engine', 'modRdv3Zip', 'modRdv3Save', 'modRdv3Be')
$allModules = @($workerModules + $modules | Select-Object -Unique)
$src = Join-Path $Root 'src\app\vba'
$destRoot = Join-Path $Root 'dist\app-vba'
$dataDir = Join-Path $destRoot 'data'
$outPath = Join-Path $destRoot 'ReaderDataViewer.xlsm'
$ledgerPath = Join-Path $destRoot 'ReaderDataViewer-Ledger.xlsx'
$sidecarPath = Join-Path $destRoot 'ReaderDataViewer-Ledger.state'

if (-not (Test-Path -LiteralPath (Join-Path $dataDir 'tableA.csv'))) {
  throw "data missing: run build_app.ps1 (it copies data-100k into $dataDir first)"
}

function Src([string] $name) {
  return (Join-Path $src "$name.bas")
}

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
"@
}

# --- preflight: name collisions, one-liners, declaration order ---------------
$declRe = '^\s*(?:Public\s+|Private\s+|Global\s+)?(?:Static\s+)?(?:Declare\s+(?:PtrSafe\s+)?)?(?:Sub|Function|Property\s+(?:Get|Let|Set)|Const|Type|Enum)\s+([A-Za-z_][A-Za-z0-9_]*)'
$varRe = '(?<![A-Za-z0-9_])([A-Za-z_][A-Za-z0-9_]*)\s*(?:\([^)]*\))?\s+As\s'
$collisions = @()
$oneLiners = @()
$misplaced = @()
foreach ($m in $allModules) {
  $seen = @{}
  $ln = 0
  $seenProc = $false
  foreach ($line in [IO.File]::ReadAllLines((Src $m))) {
    $ln++
    if ($line -match "^\s*'") { continue }
    $names = @()
    if ($line -match $declRe) { $names += $Matches[1] }
    if ($line -match '(?i)^\s*(Public|Private|Global)\s' -and
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
if ($collisions.Count) { throw ("VBA の名前衝突:`n  " + ($collisions -join "`n  ")) }
if ($oneLiners.Count) { throw ("1 行に詰めた手続き定義:`n  " + ($oneLiners -join "`n  ")) }
if ($misplaced.Count) { throw ("宣言が手続きより後ろ:`n  " + ($misplaced -join "`n  ")) }
Write-Output 'name / one-line / declaration-order preflight: ok'

# --- transcode UTF-8 -> CP932, CRLF -----------------------------------------
$stage = Join-Path $env:TEMP ("rdv3_bas_" + [Guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $stage | Out-Null
$cp932 = [Text.Encoding]::GetEncoding(932)
foreach ($m in $allModules) {
  $p = Src $m
  $t = [IO.File]::ReadAllText($p)
  if ($t -ne $cp932.GetString($cp932.GetBytes($t))) {
    throw "$p has characters that cannot survive CP932; VBE would import them as '?'"
  }
  # CRLF, always: the VBE importer is line-ending sensitive
  # (docs/results2.md trap 1)
  $t = $t -replace "`r`n", "`n"
  $t = $t -replace "`n", "`r`n"
  $ext = [IO.Path]::GetExtension($p)
  [IO.File]::WriteAllText((Join-Path $stage "$m$ext"), $t, $cp932)
}
Step "staged modules as CP932 in $stage"

# ============================================================================
# painting
# ============================================================================
$InkColor = 2500134        # RGB(38,32,28) reversed -> dark ink
$SubColor = 8210719        # muted label grey-blue
$PaperColor = 16513528     # very light paper
$PanelColor = 15987699     # light panel fill
$AccentColor = 9531969     # flat blue for buttons (BGR)
$BtnTextColor = 16777215   # white
$WarnColor = 2701998       # red-ish for the error line
$LineColor = 13421772      # light border grey

# The main screen is painted by build\ui_grid_app.ps1: the sheet is given
# cells small enough to be used as pixels, and every element of the reference is
# a merge of the block its rectangle covers. Same measured table the C# build
# lays out from, so the two screens cannot drift apart by accident.
. (Join-Path $Root 'build\ui_grid_app.ps1')

function Paint-Meta($ws) {
  $ws.Name = 'META'
  $ws.Range('B1').Value2 = 'rows (unused: ledger lives in its own workbook)'
  $ws.Range('B2').Value2 = 'updated'
  $ws.Range('B3').Value2 = 'formatver'
  $ws.Range('B4').Value2 = 'dataDir override'
  $ws.Range('B5').Value2 = 'logPath override'
  $ws.Range('B7').Value2 = 'embedded worker book (base64)'
  $ws.Range('C1').Value2 = 0
  $ws.Range('C3').Value2 = 1
}

# ============================================================================
# Excel
# ============================================================================
# identity is settled before anything is done to it (excel_own.ps1):
# a reused or unidentifiable instance throws here and is never driven
$rdvOwn = New-OwnedExcel
$xl = $rdvOwn.App
$mine = @($rdvOwn.Pid)
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AskToUpdateLinks = $false
Step ("started Excel (new pid: {0})" -f $(if ($mine.Count) { $mine -join ',' } else { 'reused' }))

try {
  # --- 1. the worker book: the BE modules, saved small -----------------------
  $workerTmp = Join-Path $stage 'rdv3worker.xlsm'
  Step 'worker book: Workbooks.Add'
  $wbW = $xl.Workbooks.Add($xlWBATWorksheet)
  $wbW.SaveAs($workerTmp, $xlsm)
  foreach ($m in $workerModules) {
    [void]$wbW.VBProject.VBComponents.Import((Join-Path $stage "$m.bas"))
    Step "  imported $m"
  }
  # the BE watches Notepad, so the worker book needs the UIA reference too
  Step '  AddFromGuid UIAutomationClient (worker book)'
  [void]$wbW.VBProject.References.AddFromGuid('{944DE083-8FB8-45CF-BCB7-C477ACB2F897}', 1, 0)
  $wbW.Save()
  Step ('  worker compile probe: ' + $xl.Run("'" + $wbW.Name + "'!modRdv3Be.Rdv3BeIsActive"))
  # the save methods carry their own arithmetic (CRC-32, its one-byte update,
  # and the deflate tables); the build checks them against the published values
  $self = $xl.Run("'" + $wbW.Name + "'!modRdv3Save.Rdv3SaveSelfTest")
  Step ('  save self-test: ' + $self)
  if ($self -notmatch 'crc\(123456789\)=CBF43926') { throw "CRC-32 self-test failed: $self" }
  if ($self -notmatch 'delta_ok=True') { throw "CRC-32 one-byte update self-test failed: $self" }
  if ($self -notmatch 'rev\(1,3\)=4' -or $self -notmatch 'lencode\(258\)=28' -or $self -notmatch 'distcode\(32768\)=29') {
    throw "deflate tables/bit order do not match RFC 1951: $self"
  }
  # ReaderDataViewer.json is read and written by BOTH builds, so the VBA half has
  # to answer what Rdv3Json/Rdv3Config answer: the same three comment forms, a
  # value out of range falling back on its own default, targets kept or dropped
  # by the same rule, and a UTF-8 round trip. Same idea as the C# side's
  # test_settings_contract.ps1, run inside the Excel that will host the code.
  $cfgSelf = $xl.Run("'" + $wbW.Name + "'!modRdv3Cfg.Rdv3CfgSelfTest")
  Step ('  settings self-test: ' + $cfgSelf)
  if ($cfgSelf -notlike 'ok *') { throw "settings self-test failed: $cfgSelf" }
  # ReaderDataViewer.json is the C# build's file too, so a VBA save must not lose
  # the members the VBA settings sheet has no column for
  $cfgRt = $xl.Run("'" + $wbW.Name + "'!modRdv3Cfg.Rdv3CfgRoundTripTest")
  Step ('  settings round trip: ' + $cfgRt)
  if ($cfgRt -notlike 'ok *') { throw "settings round trip failed: $cfgRt" }

  # seed the initial ledger workbook + sidecar THROUGH THE WORKER's own code
  # path (Rdv3BeBuildInitial = engine merge + WriteLedgerBookAll + sidecar)
  if (Test-Path -LiteralPath $ledgerPath) { Remove-Item -LiteralPath $ledgerPath -Force }
  if (Test-Path -LiteralPath $sidecarPath) { Remove-Item -LiteralPath $sidecarPath -Force }
  Step 'seed the initial ledger workbook (worker code path)'
  $seed = $xl.Run("'" + $wbW.Name + "'!modRdv3Be.Rdv3BeBuildInitial", $dataDir, $ledgerPath)
  Step "  Rdv3BeBuildInitial: $seed"
  if ($seed -notlike 'ok *') { throw "initial ledger build failed: $seed" }
  if (-not (Test-Path -LiteralPath $ledgerPath)) { throw 'ledger workbook was not written' }
  if (-not (Test-Path -LiteralPath $sidecarPath)) { throw 'ledger sidecar was not written' }

  $wbW.Save()
  $wbW.Close($false)
  $workerBytes = [IO.File]::ReadAllBytes($workerTmp)
  $workerB64 = [Convert]::ToBase64String($workerBytes)
  Step ("worker book built: {0:N0} bytes, base64 {1:N0} chars" -f $workerBytes.Length, $workerB64.Length)

  # --- 2. the main book ------------------------------------------------------
  if (Test-Path -LiteralPath $outPath) { Remove-Item -LiteralPath $outPath -Force }
  Step 'main book: Workbooks.Add'
  $wb = $xl.Workbooks.Add($xlWBATWorksheet)
  $wb.SaveAs($outPath, $xlsm)

  $wsUi = $wb.Worksheets.Item(1)
  $grid = Paint-Ui $wsUi
  Step ("  UI grid: 1 cell = {0} px (ColumnWidth {1}), {2} cols x {3} rows, card {4} px tall" -f `
        $grid.ColPx, $grid.ColWidth, $grid.Cols, $grid.Rows, [Math]::Round($grid.CardH))
  # the grid IS the layout, so a cell that is not 4 px wide is a broken screen,
  # not a cosmetic difference: say so here rather than ship it
  if ([Math]::Abs($grid.ColPx - 4.0) -gt 1.0) {
    throw ("the UI grid could not be calibrated: one cell measured {0} px, wanted 4" -f $grid.ColPx)
  }
  $wsUi.DisplayPageBreaks = $false
  # the settings screen, on the same grid. Hidden until the 設定 button asks
  # for it (modRdv3Set), so the book opens on the main screen as before.
  $wsSet = $wb.Worksheets.Add([Type]::Missing, $wsUi)
  $sg = Paint-Settings $wsSet
  Step ("  settings grid: 1 cell = {0} px, {1} cols x {2} rows" -f $sg.ColPx, $sg.Cols, $sg.Rows)

  $wsMeta = $wb.Worksheets.Add([Type]::Missing, $wsSet)
  Paint-Meta $wsMeta

  # the embedded worker book, in 30k chunks down column E
  # 8,000, not 30,000. A cell holds 32,767 characters, so the bigger chunk fits
  # on paper -- but the worker book grew with the settings and watch modules
  # (295k -> 448k base64 chars) and Excel, holding two fully painted grid sheets
  # at the same time, ran out of memory building the string for one. Smaller
  # chunks cost a few more rows in a hidden sheet and nothing else.
  $chunk = 8000
  $r = 8
  for ($i = 0; $i -lt $workerB64.Length; $i += $chunk) {
    $n = [Math]::Min($chunk, $workerB64.Length - $i)
    $wsMeta.Cells.Item($r, 5).Value2 = $workerB64.Substring($i, $n)
    $r++
  }
  Step ("embedded worker book into META (rows 8..{0})" -f ($r - 1))

  # window state: flat app look, saved with the file
  $wsUi.Activate()
  $xl.ActiveWindow.DisplayGridlines = $false
  $xl.ActiveWindow.DisplayHeadings = $false
  $wsMeta.Visible = 0     # xlSheetHidden
  $wsSet.Visible = 0      # xlSheetHidden until the 設定 button asks for it

  Step 'import modules'
  foreach ($m in $modules) {
    [void]$wb.VBProject.VBComponents.Import((Join-Path $stage "$m.bas"))
    Step "  imported $m"
  }

  # the FE needs no UIA reference: the Notepad watch lives only in the worker
  Step 'compile probe (every FE module)'
  $wb.Save()
  [void]$xl.Run("'" + (Split-Path -Leaf $outPath) + "'!modRdv3App.Rdv3BuildTouch")
  # the settings SHEET's own round trip: a C#-written configuration loaded into
  # these cells and saved again untouched must still carry the members this
  # sheet has no column for (and the targets past its last row)
  $setSelf = $xl.Run("'" + (Split-Path -Leaf $outPath) + "'!modRdv3Set.Rdv3SetSelfTest")
  Step ('  settings sheet round trip: ' + $setSelf)
  if ($setSelf -notlike 'ok *') { throw "settings sheet round trip failed: $setSelf" }
  # and the other direction: 採用 onto a row that described a DIFFERENT element
  # must leave none of the old matchers behind, or the target can never bind
  $adoptSelf = $xl.Run("'" + (Split-Path -Leaf $outPath) + "'!modRdv3Set.Rdv3SetAdoptSelfTest")
  Step ('  settings adopt round trip: ' + $adoptSelf)
  if ($adoptSelf -notlike 'ok *') { throw "settings adopt round trip failed: $adoptSelf" }

  Step 'sheet + workbook event code'
  $uiComp = $null
  $setComp = $null
  foreach ($comp in $wb.VBProject.VBComponents) {
    if ($comp.Type -eq 100 -and $comp.Name -ne 'ThisWorkbook') {
      $sheetName = $comp.Properties.Item('Name').Value
      if ($sheetName -eq 'UI') { $uiComp = $comp }
      if ($sheetName -eq 'SETTINGS') { $setComp = $comp }
    }
  }
  if ($null -eq $uiComp) { throw 'UI sheet component not found' }
  if ($null -eq $setComp) { throw 'SETTINGS sheet component not found' }
  $setComp.CodeModule.AddFromString(@'
Option Explicit

' the settings sheet's own buttons, same mechanism as the main screen
Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
    On Error Resume Next
    modRdv3Set.Rdv3SetClick Target.Range.Address(False, False)
End Sub
'@)
  $uiComp.CodeModule.AddFromString(@'
Option Explicit

' every cell button and every candidate 選択 cell is a hyperlink to itself;
' this is the single click entry point of the sheet UI
Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
    On Error Resume Next
    modRdv3App.Rdv3HandleClick Target.Range.Address(False, False)
End Sub
'@)

  $tw = $wb.VBProject.VBComponents.Item('ThisWorkbook').CodeModule
  if ($tw.CountOfLines -gt 0) { $tw.DeleteLines(1, $tw.CountOfLines) }
  $tw.AddFromString(@'
Option Explicit

' OnTime so opening returns immediately (also for a COM caller) and the update
' check starts right after, with the animation running
Private Sub Workbook_Open()
    On Error Resume Next
    Application.OnTime Now, "'" & Me.Name & "'!Rdv3AppStart"
End Sub

' UI cell writes are cosmetic and every ledger change was saved when it was
' made, so nothing real is lost by closing without a save prompt. Two things
' can still hold a close back, and they are checked in this order:
'   1. a "processed" save whose outcome is not decided yet -- the close is
'      REFUSED (nothing is torn down) and the operator is told why; it is
'      allowed again as soon as the save is confirmed or failed.
'   2. a due pump tick that cannot be canceled -- the close is DEFERRED by
'      about a second (Cancel = True) and Rdv3FinishClose closes the book once
'      the tick fired; otherwise the tick would fire against a closed book and
'      Excel would reopen it behind a security prompt.
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    If modRdv3App.Rdv3AppCloseHeldBySave() Then
        Cancel = True
        Exit Sub
    End If
    If modRdv3App.Rdv3AppPrepareClose() Then
        Me.Saved = True
    Else
        Cancel = True
    End If
End Sub

' Pump watchdog. If the OnTime chain ended without a schedule -- a tick error
' before its reschedule -- the next change or activation in THIS book re-arms
' it. These two are the workbook's own events, so they fire for this book and
' nothing else; the logic and the pump itself stay in modRdv3App. This is not
' a result-notification path: BE results arrive only through the file channel.
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error Resume Next
    modRdv3App.Rdv3PumpEnsureArmed
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    On Error Resume Next
    modRdv3App.Rdv3PumpEnsureArmed
End Sub
'@)

  $wsUi.Activate()
  $wb.Worksheets.Item('UI').Range('A1').Select() | Out-Null
  Step 'final Save'
  $wb.Save()
  $wb.Close($false)
  Step ("built {0}" -f $outPath)
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
