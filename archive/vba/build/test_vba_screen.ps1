# ============================================================================
# test_vba_screen.ps1 -- the acceptance check for the VBA build's screens.
#
#   powershell -File build\test_vba_screen.ps1
#
# The C# screens are checked by rendering them (test_ui_geometry.ps1). A
# worksheet cannot be rendered to a bitmap, so the VBA screens are checked on
# the two things that CAN be measured, and that is stated rather than implied:
#
#   1. CONTRACT   every name the VBA writes through exists in the built book.
#                 modRdv3Ui and modRdv3Set address nothing directly -- they use
#                 the builder's defined names -- so a missing name is a control
#                 that silently does nothing.
#   2. LAYOUT     no two named parts overlap, none is empty, and the grid really
#                 is the 4 px it was calibrated to. Overlap is the failure a
#                 merged-cell layout actually has: a rectangle rounded onto its
#                 neighbour's cells takes the neighbour's text with it.
#
# The workbook is opened with macros FORCE-DISABLED, so opening it does not
# start the app, does not spawn a BE and does not touch the ledger. Nothing is
# written; the book is closed without saving.
# ============================================================================
[CmdletBinding()]
param([string] $Root = "")
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'excel_own.ps1')   # exact Excel ownership, never a pid diff
[Console]::OutputEncoding = [Text.Encoding]::UTF8
if ([string]::IsNullOrEmpty($Root)) { $Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path) }

$book = Join-Path $Root 'dist\app-vba\ReaderDataViewer.xlsm'
if (-not (Test-Path -LiteralPath $book)) {
  throw "not built yet: $book (run build\build_app.ps1)"
}

# every rdv*/set* name the sources write through, taken from the sources
$src = Join-Path $Root 'src\app\vba'
$wanted = @{}
foreach ($f in @('modRdv3Ui.bas', 'modRdv3Set.bas', 'modRdv3App.bas')) {
  $t = [IO.File]::ReadAllText((Join-Path $src $f), [Text.Encoding]::UTF8)
  # a complete name only: rdvCand_ and setT are concatenation prefixes, and
  # "settings" is a log section, not a range
  foreach ($m in [regex]::Matches($t, '"((?:rdv|set)[A-Z][A-Za-z0-9_]*)"')) {
    $n = $m.Groups[1].Value
    if ($n.EndsWith('_')) { continue }
    if ($n -eq 'setT') { continue }
    $wanted[$n] = $f
  }
  # the candidate cells are built by concatenation: rdvCand_<r>_<c>
  if ($t -match 'rdvCand_') {
    for ($r = 0; $r -lt 10; $r++) { for ($c = 0; $c -lt 10; $c++) { $wanted["rdvCand_${r}_${c}"] = $f } }
  }
  # and the target rows: setT<i>_<key>
  if ($t -match 'setT') {
    foreach ($k in @('on','name','class','proc','like','fid','ftype','read','why')) {
      for ($i = 0; $i -lt 6; $i++) { $wanted["setT${i}_$k"] = $f }
    }
  }
}

# identity is settled before anything is done to it (excel_own.ps1):
# a reused or unidentifiable instance throws here and is never driven
$rdvOwn = New-OwnedExcel
$xl = $rdvOwn.App
$mine = @($rdvOwn.Pid)
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AutomationSecurity = 3            # msoAutomationSecurityForceDisable: no macros
$fail = 0
$checked = 0
try {
  $wb = $xl.Workbooks.Open($book, 0, $true)

  # ---- 1. the contract --------------------------------------------------
  $have = @{}
  foreach ($n in $wb.Names) {
    $nm = $n.Name
    if ($nm -match '^(?:.*!)?(.+)$') { $nm = $Matches[1] }
    $have[$nm] = $n
  }
  $missing = @()
  foreach ($k in ($wanted.Keys | Sort-Object)) {
    if (-not $have.ContainsKey($k)) { $missing += ("{0}  (written by {1})" -f $k, $wanted[$k]) }
  }
  if ($missing.Count -eq 0) {
    Write-Output ("  ok   contract: all {0} named parts exist" -f $wanted.Count)
  } else {
    Write-Output ("  FAIL contract: {0} names the code writes do not exist" -f $missing.Count)
    $missing | Select-Object -First 12 | ForEach-Object { Write-Output "         $_" }
    $fail++
  }

  # ---- 2. IS THERE ANYTHING ON IT ---------------------------------------
  # This check exists because the screen shipped BLANK and everything else here
  # passed. Merging the panel backgrounds made every caption and value a write
  # into somebody else's merged cell, which Excel discards without an error: the
  # names, the merges, the borders and the geometry all survived, and the
  # operator got an empty sheet. Rectangles are not a screen; text is.
  $ui = $wb.Worksheets.Item('UI')
  $filled = $xl.WorksheetFunction.CountA($ui.UsedRange)
  # 30, from the measurement: the painted screen carries 44 text cells at rest
  # (labels and captions; the readouts are empty until there is data), and the
  # screen that shipped blank carried 0
  if ($filled -ge 30) { Write-Output ("  ok   UI: {0} cells carry text" -f $filled) }
  else { Write-Output ("  FAIL UI: only {0} cells carry text -- the screen is blank" -f $filled); $fail++ }

  # and the captions a person actually reads, by name
  $captions = @{ 'rdvBtnSearch' = '検索'; 'rdvBtnClear' = '内容クリア'
                 'rdvBtnProcessed' = '処理済み'; 'rdvBtnSettings' = '設定' }
  $bad = @()
  foreach ($k in $captions.Keys) {
    $got = ''
    try { $got = [string]$ui.Range($k).Cells(1, 1).Value2 } catch { $got = '<no such name>' }
    if ($got -ne $captions[$k]) { $bad += ("{0} reads '{1}', wanted '{2}'" -f $k, $got, $captions[$k]) }
  }
  if ($bad.Count -eq 0) { Write-Output '  ok   UI: the buttons carry their captions' }
  else { $bad | ForEach-Object { Write-Output ("  FAIL " + $_) }; $fail++ }

  # ---- 3. the layout ----------------------------------------------------
  foreach ($sheetName in @('UI', 'SETTINGS')) {
    $ws = $wb.Worksheets.Item($sheetName)
    $parts = @()
    foreach ($k in $have.Keys) {
      $r = $null
      try { $r = $have[$k].RefersToRange } catch { continue }
      if ($r.Worksheet.Name -ne $sheetName) { continue }
      $parts += [pscustomobject]@{
        Name = $k
        L = [double]$r.Left; T = [double]$r.Top
        W = [double]$r.Width; H = [double]$r.Height
      }
    }
    $checked += $parts.Count
    $empty = @($parts | Where-Object { $_.W -le 0.5 -or $_.H -le 0.5 })
    if ($empty.Count -eq 0) { Write-Output ("  ok   {0}: {1} parts, none empty" -f $sheetName, $parts.Count) }
    else { Write-Output ("  FAIL {0}: {1} parts have no area" -f $sheetName, $empty.Count); $fail++ }

    # CONTAINERS are not leaves. rdvCard is the whole card -- the app fits the
    # window to it (modRdv3Ui.Rdv3UiFitWindow) -- so of course it covers every
    # part inside it. The overlap this check exists for is two LEAVES sharing
    # cells, which is what takes one part's text away with the other.
    $containers = @('rdvCard')
    $parts = @($parts | Where-Object { $containers -notcontains $_.Name })
    $hits = @()
    for ($i = 0; $i -lt $parts.Count; $i++) {
      for ($j = $i + 1; $j -lt $parts.Count; $j++) {
        $a = $parts[$i]; $b = $parts[$j]
        if ($a.L -lt $b.L + $b.W - 0.01 -and $b.L -lt $a.L + $a.W - 0.01 -and
            $a.T -lt $b.T + $b.H - 0.01 -and $b.T -lt $a.T + $a.H - 0.01) {
          $hits += ("{0} x {1}" -f $a.Name, $b.Name)
        }
      }
    }
    if ($hits.Count -eq 0) { Write-Output ("  ok   {0}: no two parts overlap" -f $sheetName) }
    else {
      Write-Output ("  FAIL {0}: {1} overlapping pairs" -f $sheetName, $hits.Count)
      $hits | Select-Object -First 10 | ForEach-Object { Write-Output "         $_" }
      $fail++
    }

    # the grid unit itself, in px (Excel reports points)
    $px = [double]$ws.Columns.Item(3).Width / 0.75
    if ([Math]::Abs($px - 4.0) -le 1.0) { Write-Output ("  ok   {0}: 1 cell = {1:N2} px" -f $sheetName, $px) }
    else { Write-Output ("  FAIL {0}: 1 cell = {1:N2} px, wanted 4" -f $sheetName, $px); $fail++ }
  }

  $wb.Close($false)
}
finally {
  try { $xl.Quit() } catch { }
  try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) } catch { }
  # only the instance this test created, and only if Quit did not take
  foreach ($id in $mine) { Stop-ExcelOwned $id }
}

Write-Output ''
Write-Output ("{0} named parts measured" -f $checked)
if ($fail -gt 0) { Write-Output 'RESULT: FAIL'; exit 1 }
Write-Output 'RESULT: PASS'
