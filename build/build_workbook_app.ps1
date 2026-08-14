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

$modules = @('modRdv3Spec', 'modRdv3Chan', 'modRdv3Host', 'modRdv3Ui', 'modRdv3App')
$classes = @('clsRdv3AppEvents')
$workerModules = @('modRdv3Spec', 'modRdv3Chan', 'modRdv3Uia', 'modRdv3Engine', 'modRdv3Be')
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
  if ($classes -contains $name) { return (Join-Path $src "$name.cls") }
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
foreach ($m in ($allModules + $classes)) {
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
foreach ($m in ($allModules + $classes)) {
  $p = Src $m
  $t = [IO.File]::ReadAllText($p)
  if ($t -ne $cp932.GetString($cp932.GetBytes($t))) {
    throw "$p has characters that cannot survive CP932; VBE would import them as '?'"
  }
  # CRLF, always: VBE only recognises the VERSION/BEGIN/END header of a .cls
  # when the lines end in CRLF (docs/results2.md trap 1)
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

function Add-CellButton($ws, [string] $addr, [string] $caption, [string] $tip) {
  $cell = $ws.Range($addr)
  $cell.Value2 = $caption
  [void]$ws.Hyperlinks.Add($cell, "", ($ws.Name + "!" + $addr), $tip, $caption)
  # Hyperlinks.Add re-styles the cell; make it a flat button again
  $cell.Font.Name = 'Yu Gothic UI'
  $cell.Font.Size = 10
  $cell.Font.Bold = $true
  $cell.Font.Underline = $false
  $cell.Font.Color = $BtnTextColor
  $cell.Interior.Color = $AccentColor
  $cell.HorizontalAlignment = -4108   # center
  $cell.VerticalAlignment = -4108
  $cell.Borders.LineStyle = 1
  $cell.Borders.Color = $AccentColor
}

function Paint-Ui($ws) {
  $ws.Name = 'UI'
  $ws.Cells.Font.Name = 'Yu Gothic UI'
  $ws.Cells.Font.Size = 10
  $ws.Cells.Interior.Color = $PaperColor

  $ws.Columns('A').ColumnWidth = 2
  $ws.Columns('B').ColumnWidth = 12
  $ws.Columns('C').ColumnWidth = 20
  $ws.Columns('D').ColumnWidth = 12
  $ws.Columns('E').ColumnWidth = 20
  $ws.Columns('F').ColumnWidth = 14
  $ws.Columns('G').ColumnWidth = 20
  $ws.Columns('H').ColumnWidth = 20
  $ws.Columns('I').ColumnWidth = 13
  $ws.Columns('J').ColumnWidth = 13
  $ws.Columns('K').ColumnWidth = 2
  $ws.Rows.Item(1).RowHeight = 30

  $ws.Range('B1').Value2 = 'Reader Data Viewer'
  $ws.Range('B1').Font.Size = 18
  $ws.Range('B1').Font.Bold = $true
  $ws.Range('E1').Value2 = 'VBA 標準 Dictionary'
  $ws.Range('E1').Font.Size = 12
  $ws.Range('E1').Font.Bold = $true
  $ws.Range('E1').Font.Color = $SubColor

  # ---- status ----
  $ws.Range('B3').Value2 = '状態'
  $ws.Range('B4').Value2 = 'メモ帳'
  $ws.Range('B5').Value2 = '台帳'
  $ws.Range('B6').Value2 = 'マージ時間'
  $ws.Range('D6').Value2 = '検索時間'
  $ws.Range('B7').Value2 = 'エラー'
  $ws.Range('B3:B7').Font.Color = $SubColor
  $ws.Range('D6').Font.Color = $SubColor
  $ws.Range('C3').Font.Bold = $true
  $ws.Range('C3').Font.Size = 12
  $ws.Range('C6').Font.Name = 'Consolas'
  $ws.Range('C6').Font.Bold = $true
  $ws.Range('E6').Font.Name = 'Consolas'
  $ws.Range('E6').Font.Bold = $true
  $ws.Range('C7').Font.Color = $WarnColor
  $ws.Range('C7').Font.Bold = $true
  $ws.Range('B2:J7').Interior.Color = 16777215
  $b = $ws.Range('B2:J7').Borders
  $b.LineStyle = 1
  $b.Color = $LineColor
  $ws.Range('B2:J7').Borders.Item(11).LineStyle = -4142   # no inner vertical
  $ws.Range('B2:J7').Borders.Item(12).LineStyle = -4142   # no inner horizontal

  # ---- operations ----
  $ws.Range('B9').Value2 = '検索番号'
  $ws.Range('B9').Font.Color = $SubColor
  $c = $ws.Range('C9')
  $c.NumberFormat = '@'
  $c.Interior.Color = 16777215
  $c.Font.Name = 'Consolas'
  $c.Font.Size = 12
  $c.HorizontalAlignment = -4108
  $c.Borders.LineStyle = 1
  $c.Borders.Color = $SubColor
  [void]$c.Validation.Add(0)   # xlValidateInputOnly: message only, no rule
  $c.Validation.InputTitle = '番号1'
  $c.Validation.InputMessage = "8 桁の数字を入力し、「検索」をクリックしてください。`nメモ帳の自動検知でも検索できます。"
  $c.Validation.ShowInput = $true

  Add-CellButton $ws 'E9' '検索' '入力した番号1 で統合台帳を検索します'
  Add-CellButton $ws 'F9' '内容クリア' '入力と結果表示を消します (台帳と処理済みは消えません)'
  Add-CellButton $ws 'G9' '処理済み' '表示中の統合レコードを処理済みにします (確認あり)'
  Add-CellButton $ws 'I9' '監視停止' 'メモ帳の自動検知を止める / 再開する'

  # ---- current key / verdict ----
  $ws.Range('B11').Value2 = '現在の番号 (番号1)'
  $ws.Range('B11').Font.Color = $SubColor
  $ws.Range('C11').Font.Name = 'Consolas'
  $ws.Range('C11').Font.Size = 16
  $ws.Range('C11').Font.Bold = $true
  $ws.Range('C11').NumberFormat = '@'
  $ws.Range('E11').Font.Size = 11
  $ws.Range('E11').Font.Bold = $true
  $ws.Range('I11').Font.Bold = $true
  $ws.Range('B11:J11').Interior.Color = 15854037

  # ---- record ----
  $hdr = @('#', '表A 項目', '表A 値', '表B 項目', '表B 値', '表C 項目', '表C 値')
  for ($i = 0; $i -lt $hdr.Count; $i++) {
    $cell = $ws.Cells.Item(13, 2 + $i)
    $cell.Value2 = $hdr[$i]
    $cell.Font.Bold = $true
    $cell.Interior.Color = 14994616
  }
  $ws.Range('B14:H23').NumberFormat = '@'
  $ws.Range('B14:H23').Interior.Color = 16777215
  $ws.Range('B14:B23').HorizontalAlignment = -4152
  $ws.Range('B13:H23').Borders.LineStyle = 1
  $ws.Range('B13:H23').Borders.Color = $LineColor

  # ---- candidates ----
  $ws.Range('B25').Value2 = '候補一覧 (行頭の「選択」で表示)'
  $ws.Range('B25').Font.Color = $SubColor
  $h3 = @('選択', '番号2', '行番号', '伝票番号', '日付', '数量', '状態', '品目コード', 'メーカー')
  for ($i = 0; $i -lt $h3.Count; $i++) {
    $cell = $ws.Cells.Item(26, 2 + $i)
    $cell.Value2 = $h3[$i]
    $cell.Font.Bold = $true
    $cell.Interior.Color = 14994616
  }
  $ws.Range('B27:J36').NumberFormat = '@'
  $ws.Range('B27:J36').Interior.Color = 16777215
  $ws.Range('B26:J36').Borders.LineStyle = 1
  $ws.Range('B26:J36').Borders.Color = $LineColor
  for ($r = 27; $r -le 36; $r++) {
    $cell = $ws.Cells.Item($r, 2)
    [void]$ws.Hyperlinks.Add($cell, "", ("UI!B" + $r), 'この候補の統合レコードを表示します', ' ')
    $cell.Value2 = ''
    $cell.Font.Name = 'Yu Gothic UI'
    $cell.Font.Size = 10
    $cell.Font.Bold = $true
    $cell.Font.Underline = $false
    $cell.Font.Color = $AccentColor
    $cell.HorizontalAlignment = -4108
  }

  $ws.Range('B38').Value2 = 'メモ帳に 8 桁の番号を入力すると自動で検索します。手動のときは検索番号に入力して「検索」。'
  $ws.Range('B39').Value2 = '複数候補のときは一覧の「選択」をクリックすると、その統合レコードを表示します。'
  $ws.Range('B38:B39').Font.Color = $SubColor
  $ws.Range('A1').Select() | Out-Null
}

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
$before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AskToUpdateLinks = $false
$after = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$mine = @($after | Where-Object { $before -notcontains $_ })
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
  Paint-Ui $wsUi
  $wsMeta = $wb.Worksheets.Add([Type]::Missing, $wsUi)
  Paint-Meta $wsMeta

  # the embedded worker book, in 30k chunks down column E
  $chunk = 30000
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

  Step 'import modules'
  foreach ($m in $modules) {
    [void]$wb.VBProject.VBComponents.Import((Join-Path $stage "$m.bas"))
    Step "  imported $m"
  }
  foreach ($m in $classes) {
    [void]$wb.VBProject.VBComponents.Import((Join-Path $stage "$m.cls"))
    Step "  imported $m (class)"
  }

  # the FE needs no UIA reference: the Notepad watch lives only in the worker
  Step 'compile probe (every FE module + the class)'
  $wb.Save()
  [void]$xl.Run("'" + (Split-Path -Leaf $outPath) + "'!modRdv3App.Rdv3BuildTouch")

  Step 'sheet + workbook event code'
  $uiComp = $null
  foreach ($comp in $wb.VBProject.VBComponents) {
    if ($comp.Type -eq 100 -and $comp.Name -ne 'ThisWorkbook') {
      if ($comp.Properties.Item('Name').Value -eq 'UI') { $uiComp = $comp }
    }
  }
  if ($null -eq $uiComp) { throw 'UI sheet component not found' }
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
' made, so nothing real is lost by closing without a save prompt. When a due
' pump tick cannot be canceled the close is deferred by about a second
' (Cancel = True) and Rdv3FinishClose closes the book once the tick fired --
' otherwise the tick would fire against a closed book and Excel would reopen
' it behind a security prompt.
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    If modRdv3App.Rdv3AppPrepareClose() Then
        Me.Saved = True
    Else
        Cancel = True
    End If
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
