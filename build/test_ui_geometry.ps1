# ============================================================================
# test_ui_geometry.ps1 -- the acceptance check for the practical build's screen.
#
#   powershell -File build\test_ui_geometry.ps1
#   powershell -File build\test_ui_geometry.ps1 -Quick
#   powershell -File build\test_ui_geometry.ps1 -Png        (also writes bitmaps)
#
# The screen was rebuilt to the UI reference (Reader Data Viewer_ver3.html, and
# docs\ui-spec.md). "It looks like the picture" is not the condition; the
# conditions are:
#
#   1. FIDELITY   at the design size, every named element sits within -Tol of
#                 the rectangle measured from the reference (docs\ui-ref-geom.json).
#                 Layout decides y/h for every element and x/w for the
#                 structural ones; the rest follow their text, so a substituted
#                 font may legitimately move them (docs\ui-spec.md 8.1).
#   2. INTEGRITY  at EVERY size, scaling factor, state and data set: no two leaf
#                 rectangles intersect, nothing falls outside the client area,
#                 and no string is silently cut off. This is the condition a
#                 screen that merely imitates the picture fails.
#
# The form is compiled from the SHIPPING sources and rendered WITHOUT EVER
# BEING SHOWN (the handle is created, DrawToBitmap paints it). No window
# appears, the foreground never moves, and what is measured is the product.
# ============================================================================
[CmdletBinding()]
param(
  [string] $Root = "",
  [string] $Out = "",
  [double[]] $Scales = @(1.0, 1.25, 1.5),
  [string[]] $Sizes = @('1240x689', '1366x768', '1480x800', '1920x1040'),
  [string[]] $States = @('boot', 'checking', 'applying', 'watching', 'waiting', 'searching',
                         'none', 'single', 'multi', 'picked', 'processed', 'saving', 'error', 'cleared',
                         'badkey'),
  [string[]] $Data = @('ref', 'long'),
  [double] $Tol = 2.0,
  [switch] $Quick,
  [switch] $Png
)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
if ([string]::IsNullOrEmpty($Root)) { $Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path) }
if ([string]::IsNullOrEmpty($Out)) { $Out = Join-Path $Root 'work\ui-check' }
if ($Quick) { $Scales = @(1.0); $Sizes = @('1240x689'); $States = @('multi', 'picked', 'error'); $Data = @('ref') }

# ---- the product, compiled exactly as the packer compiles it ---------------
$sources = @('Rdv3Core.cs', 'Rdv3Index.cs', 'Rdv3Ledger.cs', 'Rdv3Xlsx.cs', 'Rdv3Jobs.cs',
             'Rdv3Watch.cs', 'Rdv3Text.cs', 'Rdv3Geom.cs', 'Rdv3SetGeom.cs', 'Rdv3Json.cs', 'Rdv3Config.cs',
             'Rdv3Uia.cs', 'Rdv3Ui.cs', 'Rdv3Settings.cs', 'Rdv3App.cs')
$usings = New-Object System.Collections.Specialized.OrderedDictionary
$bodies = New-Object System.Text.StringBuilder
foreach ($f in $sources) {
  $t = [IO.File]::ReadAllText((Join-Path $Root "src\app\csharp\$f"), [Text.Encoding]::UTF8)
  if ($t.Contains('@"')) { throw "$f has a verbatim string" }
  foreach ($line in ($t -split "`r?`n")) {
    if ($line -match '^\s*using\s+[A-Za-z_][A-Za-z0-9_.]*\s*;\s*$') {
      $k = $line.Trim(); if (-not $usings.Contains($k)) { $usings.Add($k, $true) }
    } else { [void]$bodies.AppendLine($line) }
  }
}
$cs = (($usings.Keys | ForEach-Object { $_ }) -join "`r`n") + "`r`n`r`n" + $bodies.ToString()
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.Xml
$refs = @(
  [System.Diagnostics.Process].Assembly.Location,
  [System.Windows.Forms.Form].Assembly.Location,
  [System.Drawing.Point].Assembly.Location,
  [System.Windows.Automation.AutomationElement].Assembly.Location,
  [System.Windows.Automation.AutomationElementIdentifiers].Assembly.Location,
  [System.Windows.DependencyObject].Assembly.Location,
  [System.IO.Compression.ZipArchive].Assembly.Location,
  [System.Xml.XmlReader].Assembly.Location)
try {
  Add-Type -TypeDefinition $cs -ReferencedAssemblies $refs -Language CSharp
} catch {
  Write-Output 'COMPILE FAILED'
  Write-Output $_.Exception.Message
  if ($_.Exception.InnerException) { Write-Output $_.Exception.InnerException.Message }
  exit 1
}
Write-Output 'compile ok'
[void][System.IO.Directory]::CreateDirectory($Out)

# ---- sample data -----------------------------------------------------------
# 30 tab separated ledger fields; the screen reads 1 key2, 2..9 table A,
# 16 status, 17 line, 18 memo, 19..20 item/maker, 27 remark
function New-Line([string] $key1, [string] $key2, [hashtable] $o) {
  $f = New-Object 'string[]' 30
  for ($i = 0; $i -lt 30; $i++) { $f[$i] = '' }
  $f[0] = $key1; $f[1] = $key2
  $f[2] = $o.code; $f[3] = $o.name; $f[4] = $o.grade; $f[5] = $o.date; $f[6] = $o.amount
  $f[7] = $o.rate; $f[8] = $o.flag; $f[9] = $o.dept
  $f[11] = $o.slip; $f[12] = $o.bdate; $f[13] = $o.qty; $f[16] = $o.status
  $f[17] = $o.line; $f[18] = $o.memo; $f[19] = $o.item; $f[20] = $o.maker; $f[27] = $o.remark
  return ($f -join "`t")
}
# 'ref' repeats the artifact's own sample strings, so the design-size comparison
# is like for like; 'long' is deliberately longer than anything real
$refRow = @{ code = 'A-4471'; name = '相模原精密工業 株式会社'; grade = 'B'; date = '2024-11-08';
  amount = '1,284,000'; rate = '0.86'; flag = 'Y'; dept = '第二製造部'; slip = 'SL00016168';
  bdate = '20230117'; qty = '597'; status = 'HOLD'; line = '001'; item = 'IT347201'; maker = 'MAKER-2758';
  memo = '検品時に外観キズ 2 件。再検査済み、出荷可否は品証判断待ち。'; remark = 'ロット L-2291 の在庫と突合済み。次回入荷は 2026-09-02 予定。' }
$longRow = @{ code = 'A-4471-EXTENDED-CODE'; name = 'とても長い取引先名称株式会社 東日本第二営業統括本部'; grade = 'SPECIAL';
  date = '2024-11-08'; amount = '9,999,999,999'; rate = '0.8642'; flag = 'YES-LONG'; dept = '第二製造部 品質保証課 検査係';
  slip = 'SL-882140-EXT-0001'; bdate = '2026-08-14'; qty = '120,000'; status = 'HOLD'; line = '11482';
  item = 'IT-77120-LONG-ITEM'; maker = 'KYOWA-INDUSTRIAL-LONG';
  memo = ('長い摘要テキスト。' * 12); remark = ('長い備考テキスト。' * 12) }

# a real ledger row whose table A side never matched: every field the record
# panel shows on the left is empty, which is what N/A is for
$sparseRow = @{ code = ''; name = ''; grade = ''; date = ''; amount = ''; rate = ''; flag = ''; dept = '';
  slip = 'SL00016168'; bdate = '20230117'; qty = '597'; status = 'OPEN'; line = '004'; item = 'IT347201';
  maker = 'MAKER-2758'; memo = ''; remark = '' }

function New-Rows([int] $n, [hashtable] $o, [string] $key1) {
  $list = New-Object 'System.Collections.Generic.List[Rdv3CandRow]'
  for ($i = 0; $i -lt $n; $i++) {
    $line = New-Line $key1 ('{0:D8}' -f (60919 + $i + 1)) $o
    $r = New-Object Rdv3CandRow
    $r.Cols = [Rdv3Ledger]::CandidateColumns($line)
    $r.Processed = ($i -eq 1)
    $r.Line = $line
    [void]$list.Add($r)
  }
  return $list
}

# ---- the form, never shown -------------------------------------------------
$form = New-Object Rdv3Form
$null = $form.Handle              # creates the window handle WITHOUT showing it
$form.Visible = $false
# a form that is never shown does not create its children's handles, and
# DrawToBitmap would then render empty boxes where the buttons, the search box
# and the candidate table are
function Touch-Handles($c) {
  $null = $c.Handle
  foreach ($k in $c.Controls) { Touch-Handles $k }
}
Touch-Handles $form

function Set-Data([string] $kind) {
  if ($kind -eq 'long') {
    $form.SetIdentity('とても長い利用者名 太郎', 'DESKTOP-LONGHOSTNAME-01', '読み取り専用', 'PID 182332', 'ReaderDataViewer-very-long-name.log')
    $form.SetLedgerInfo('ReaderDataViewer-Ledger-very-long-file-name.xlsx ・ 1,284,900 件')
    $form.SetLedgerSummary('1,284,900', '08-16 01:14')
    $form.SetNotepad('とても長いタイトルのテキストファイル - メモ帳（hwnd 132850）')
    $form.SetMergeMs(41234.6); $form.SetSearchMs(1234.8)
  } else {
    $form.SetIdentity('田中 太郎', '検品担当 ・ D613', '一般', 'PID 18332', 'ReaderDataViewer.log')
    $form.SetLedgerInfo('ReaderDataViewer-Ledger.xlsx ・ 100,000 件')
    $form.SetLedgerSummary('100,000', '08-15 09:12')
    $form.SetNotepad('無題 - メモ帳（hwnd 132850）')
    $form.SetMergeMs(412.6); $form.SetSearchMs(0.8)
  }
}

function Set-State([string] $st, [string] $kind) {
  $o = if ($kind -eq 'long') { $longRow } elseif ($kind -eq 'sparse') { $sparseRow } else { $refRow }
  $key1 = if ($kind -eq 'long') { '40218764' } else { '00016168' }
  if ($kind -eq 'sparse') { $kind = 'ref' }
  $form.SetError('')
  $form.ClearResult()
  Set-Data $kind
  switch ($st) {
    'boot'      { $form.SetState([Rdv3Text]::StateBoot, 0) }
    'checking'  { $form.SetState([Rdv3Text]::StateChecking, 0) }
    'applying'  { $form.SetState([Rdv3Text]::StateApplying, 0) }
    'watching'  { $form.SetState([Rdv3Text]::StateReady, 0) }
    'waiting'   { $form.SetState([Rdv3Text]::StateWaitingNotepad, 0); $form.SetNotepad([Rdv3Text]::NotepadNone) }
    'searching' { $form.SetState([Rdv3Text]::StateBusy, 0); $form.SetKeyText($key1) }
    'none'      { $form.SetState([Rdv3Text]::StateReady, 0); $form.SetKeyText($key1)
                  $form.ShowCandidates($key1, (New-Rows 0 $o $key1), 0, 0) }
    'single'    { $form.SetState([Rdv3Text]::StateReady, 0); $form.SetKeyText($key1)
                  $rows = New-Rows 1 $o $key1
                  $form.ShowCandidates($key1, $rows, 1, 1)
                  $form.SelectCandidate(0, $rows[0].Line, $false, 1) }
    'multi'     { $form.SetState([Rdv3Text]::StateReady, 0); $form.SetKeyText($key1)
                  $form.ShowCandidates($key1, (New-Rows 8 $o $key1), 8, 8) }
    'multi3'    { $form.SetState([Rdv3Text]::StateReady, 0); $form.SetKeyText($key1)
                  $form.ShowCandidates($key1, (New-Rows 3 $o $key1), 3, 3)
                  $form.SetSearchMs(7.33) }
    'picked'    { $form.SetState([Rdv3Text]::StateReady, 0); $form.SetKeyText($key1)
                  $rows = New-Rows 8 $o $key1
                  $form.ShowCandidates($key1, $rows, 8, 8)
                  $form.SelectCandidate(2, $rows[2].Line, $false, 8) }
    'processed' { $form.SetState([Rdv3Text]::StateReady, 0); $form.SetKeyText($key1)
                  $rows = New-Rows 8 $o $key1
                  $form.ShowCandidates($key1, $rows, 8, 8)
                  $form.SelectCandidate(2, $rows[2].Line, $true, 8)
                  $form.ShowProcessedState('TRUE') }
    'saving'    { $form.SetState([Rdv3Text]::StateSavingMark, 0); $form.SetKeyText($key1)
                  $rows = New-Rows 8 $o $key1
                  $form.ShowCandidates($key1, $rows, 8, 8)
                  $form.SelectCandidate(2, $rows[2].Line, $false, 8)
                  $form.ShowProcessedState('TRUE' + [Rdv3Text]::SavingSuffix) }
    'error'     { $form.SetState([Rdv3Text]::StateReady, 0)
                  $rows = New-Rows 8 $o $key1
                  $form.ShowCandidates($key1, $rows, 8, 8)
                  $form.SelectCandidate(2, $rows[2].Line, $false, 8)
                  $form.SetError([Rdv3Text]::ErrPersist + 'ReaderDataViewer-Ledger.xlsx は他のプロセスが使用中です') }
    'cleared'   { $form.SetState([Rdv3Text]::StateReady, 0); $form.ClearResult() }
    'badkey'    { $form.SetState([Rdv3Text]::StateReady, 0); $form.SetKeyText('123')
                  $form.SetInputError([Rdv3Text]::ErrBadKey.Replace('{n}', '8')) }
    default     { throw "unknown state: $st" }
  }
  [System.Windows.Forms.Application]::DoEvents()
}

# ---- the two checks --------------------------------------------------------
# containers legitimately hold other rectangles, so they are not paired
$containers = @('nav', 'sum.panel', 'list.panel', 'list.scroll', 'rec.panel', 'status', 'err',
                'nav.rule', 'list.rule', 'rec.rule')
# x/w of these is fixed by the layout; for everything else it follows the text
$structural = @('nav', 'sum.panel', 'list.panel', 'list.scroll', 'rec.panel', 'status', 'status.dot',
                'rec.memoLabel', 'rec.memoBox', 'rec.remarkLabel', 'rec.remarkBox')
for ($i = 0; $i -lt 7; $i++) { $structural += ('rec.kv{0}' -f $i) }
for ($i = 0; $i -lt 10; $i++) { $structural += ('list.col{0}' -f $i); $containers += ('list.col{0}' -f $i) }
# the input keeps the reference's SIZE exactly but sits at the right hand end of
# the summary row, so its x follows the width of the text blocks beside it
$rightAnchored = @('sum.input', 'sum.field.label')

function Test-Integrity($el, [double] $cw, [double] $ch) {
  $keys = @($el.PSObject.Properties.Name | Where-Object { $containers -notcontains $_ })
  $n = $keys.Count
  $x0 = New-Object 'double[]' $n; $y0 = New-Object 'double[]' $n
  $x1 = New-Object 'double[]' $n; $y1 = New-Object 'double[]' $n
  for ($i = 0; $i -lt $n; $i++) {
    $r = $el.($keys[$i])
    $x0[$i] = $r[0]; $y0[$i] = $r[1]; $x1[$i] = $r[0] + $r[2]; $y1[$i] = $r[1] + $r[3]
  }
  # sweep in y: only rectangles whose vertical ranges overlap can intersect
  $order = 0..($n - 1) | Sort-Object { $y0[$_] }
  $hits = New-Object System.Collections.ArrayList
  for ($a = 0; $a -lt $n; $a++) {
    $i = $order[$a]
    for ($b = $a + 1; $b -lt $n; $b++) {
      $j = $order[$b]
      if ($y0[$j] -ge $y1[$i] - 0.6) { break }
      $w = [Math]::Min($x1[$i], $x1[$j]) - [Math]::Max($x0[$i], $x0[$j])
      $h = [Math]::Min($y1[$i], $y1[$j]) - [Math]::Max($y0[$i], $y0[$j])
      if ($w -gt 0.6 -and $h -gt 0.6) {
        [void]$hits.Add(('{0} x {1} ({2:F1}x{3:F1})' -f $keys[$i], $keys[$j], $w, $h))
      }
    }
  }
  $outside = New-Object System.Collections.ArrayList
  foreach ($k in $el.PSObject.Properties.Name) {
    $r = $el.$k
    if ($r[0] -lt -0.6 -or $r[1] -lt -0.6 -or ($r[0] + $r[2]) -gt ($cw + 0.6) -or ($r[1] + $r[3]) -gt ($ch + 0.6)) {
      [void]$outside.Add(('{0} [{1},{2},{3},{4}]' -f $k, $r[0], $r[1], $r[2], $r[3]))
    }
  }
  return @{ Overlap = $hits; Outside = $outside }
}

$refPath = Join-Path $Root 'docs\ui-ref-geom.json'
$refGeom = (Get-Content -LiteralPath $refPath -Raw -Encoding UTF8 | ConvertFrom-Json).el

function Test-Fidelity($el) {
  $rows = New-Object System.Collections.ArrayList
  $bad = 0; $missing = 0
  foreach ($k in ($refGeom.PSObject.Properties.Name | Sort-Object)) {
    $r = $refGeom.$k
    $a = $el.$k
    if ($null -eq $a) { $missing++; continue }
    $d = @(($a[0] - $r[0]), ($a[1] - $r[1]), ($a[2] - $r[2]), ($a[3] - $r[3]))
    if ($structural -contains $k) { $cls = 'geom'; $chk = @(0, 1, 2, 3) }
    elseif ($rightAnchored -contains $k) { $cls = 'size'; $chk = @(1, 2, 3) }
    else { $cls = 'text'; $chk = @(1, 3) }
    $worst = 0.0
    foreach ($c in $chk) { if ([Math]::Abs($d[$c]) -gt $worst) { $worst = [Math]::Abs($d[$c]) } }
    if ($worst -gt $Tol) { $bad++ }
    [void]$rows.Add([pscustomobject]@{ Element = $k; Class = $cls
      Reference = ('{0},{1},{2},{3}' -f $r[0], $r[1], $r[2], $r[3])
      App = ('{0},{1},{2},{3}' -f $a[0], $a[1], $a[2], $a[3])
      Delta = ('{0},{1},{2},{3}' -f $d[0], $d[1], $d[2], $d[3])
      Worst = [Math]::Round($worst, 1); Over = ($worst -gt $Tol) })
  }
  return @{ Rows = $rows; Bad = $bad; Missing = $missing }
}

# ---- run -------------------------------------------------------------------
$cases = 0; $failed = 0
$baseEl = $null
$sw = [Diagnostics.Stopwatch]::StartNew()
foreach ($sc in $Scales) {
  $form.SetScale([single]$sc)
  foreach ($sz in $Sizes) {
    $wh = $sz -split 'x'
    $form.ClientSize = New-Object Drawing.Size ([int]([double]$wh[0] * $sc)), ([int]([double]$wh[1] * $sc))
    foreach ($kind in $Data) {
      foreach ($st in $States) {
        Set-State $st $kind
        $tag = '{0}-{1}-{2}-{3}' -f $st, $kind, $sz, ($sc.ToString('0.00'))
        # render first: the clipping report is filled while painting
        $bmp = New-Object Drawing.Bitmap $form.Width, $form.Height
        $form.DrawToBitmap($bmp, (New-Object Drawing.Rectangle 0, 0, $form.Width, $form.Height))
        $json = $form.GeometryDump()
        [IO.File]::WriteAllText((Join-Path $Out ("geom-$tag.json")), $json, (New-Object Text.UTF8Encoding $false))
        if ($Png) {
          $o = $form.PointToScreen([Drawing.Point]::Empty)
          $crop = New-Object Drawing.Rectangle ($o.X - $form.Left), ($o.Y - $form.Top), $form.ClientSize.Width, $form.ClientSize.Height
          $c2 = $bmp.Clone($crop, $bmp.PixelFormat)
          $c2.Save((Join-Path $Out ("shot-$tag.png")), [Drawing.Imaging.ImageFormat]::Png)
          $c2.Dispose()
        }
        $bmp.Dispose()

        $d = $json | ConvertFrom-Json
        $r = Test-Integrity $d.el $d.client[0] $d.client[1]
        $clip = @($d.clipped)
        $cases++
        if ($r.Overlap.Count -gt 0 -or $r.Outside.Count -gt 0 -or $clip.Count -gt 0) {
          $failed++
          Write-Output ('FAIL {0,-40} overlap {1} outside {2} clipped {3}' -f $tag, $r.Overlap.Count, $r.Outside.Count, $clip.Count)
          foreach ($h in $r.Overlap) { Write-Output ("       overlap  " + $h) }
          foreach ($h in $r.Outside) { Write-Output ("       outside  " + $h) }
          if ($clip.Count -gt 0) { Write-Output ("       clipped  " + ($clip -join ', ')) }
        }
        if ($st -eq 'picked' -and $kind -eq 'ref' -and $sz -eq '1240x689' -and $sc -eq 1.0) { $baseEl = $d.el }
      }
    }
  }
}
$form.Dispose()

Write-Output ''
Write-Output ('integrity : {0} cases, {1} failing   ({2:F0} s)' -f $cases, $failed, $sw.Elapsed.TotalSeconds)

$fidBad = 0
if ($null -ne $baseEl) {
  $f = Test-Fidelity $baseEl
  $fidBad = $f.Bad + $f.Missing
  Write-Output ''
  Write-Output ("fidelity at the design size (tolerance {0} px)" -f $Tol)
  Write-Output ('{0,-20} {1,-5} {2,-26} {3,-26} {4}' -f 'element', 'class', 'reference x,y,w,h', 'app x,y,w,h', 'delta')
  foreach ($row in $f.Rows) {
    Write-Output ('{0,-20} {1,-5} {2,-26} {3,-26} {4}{5}' -f $row.Element, $row.Class, $row.Reference, $row.App, $row.Delta,
      $(if ($row.Over) { '   <-- over' } else { '' }))
  }
  Write-Output ('elements {0}, outside tolerance {1}, missing {2}' -f $f.Rows.Count, $f.Bad, $f.Missing)
} else {
  Write-Output 'fidelity  : the design-size case was not in this run (use the default arguments)'
}

Write-Output ''
if ($failed -eq 0 -and $fidBad -eq 0) { Write-Output 'RESULT: PASS'; exit 0 }
Write-Output ('RESULT: FAIL (integrity {0}, fidelity {1})' -f $failed, $fidBad)
exit 1
