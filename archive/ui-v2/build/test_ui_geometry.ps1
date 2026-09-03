# ============================================================================
# test_ui_geometry.ps1 -- the acceptance check for the main screen.
#
#   powershell -File build\test_ui_geometry.ps1
#   powershell -File build\test_ui_geometry.ps1 -Quick
#   powershell -File build\test_ui_geometry.ps1 -Png        (also writes bitmaps)
#   powershell -File build\test_ui_geometry.ps1 -Settings samples\sales-wide\settings.json
#
# The screen is built from the "screen" part of src\config\settings.json to the v2
# reference (docs\ui-reference\v2.html, measured into docs\ui-ref-v2-geom.json).
# With -Settings the screen of ANY definition is built instead: its sample
# rows are made up from its own ledger columns (dates and numbers in the
# shape its formats ask for, a value that judges ok / ng / undefined found
# from its own rules), and only the INTEGRITY check runs -- the fidelity
# table belongs to the shipped definition alone.
# "It looks like the picture" is not the condition; the conditions are:
#
#   1. FIDELITY   at the design size (1240 wide, 100%), every named element
#                 sits within -Tol of the rectangle measured from the reference.
#                 Structural elements are held on x/y/w/h; elements that
#                 follow their text (figures, buttons, the band's group) on y/h.
#   2. INTEGRITY  at EVERY size, scaling factor, state and data set: no two
#                 leaf rectangles intersect, nothing falls outside the client
#                 area, and no string is silently cut off (an ellipsis where
#                 the design allows one -- long values, the bar's segments --
#                 is not a cut).
#
# The form is compiled from the SHIPPING sources and rendered WITHOUT EVER
# BEING SHOWN (the handle is created, DrawToBitmap paints it). The sizes
# include ones narrower and shorter than the design, which is where the
# responsive rules (stacked columns, wrapped input group, compressed rows,
# scrolling) have to hold.
# ============================================================================
[CmdletBinding()]
param(
  [string] $Root = "",
  [string] $Out = "",
  [double[]] $Scales = @(1.0, 1.25, 1.5),
  [string[]] $Sizes = @('1240x974', '1000x700', '1480x900', '1920x1040', '800x600', '640x900'),
  [string[]] $States = @('idle', 'checking', 'single-ok', 'single-ng', 'multi', 'picked', 'saving',
                         'none', 'unknown-state', 'undefined', 'cleared'),
  [string[]] $Data = @('ref', 'long'),
  [double] $Tol = 2.0,
  [string] $Settings = "",
  [switch] $Quick,
  [switch] $Png
)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
if ([string]::IsNullOrEmpty($Root)) { $Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path) }
if ([string]::IsNullOrEmpty($Out)) { $Out = Join-Path $Root 'work\ui-check' }
if ($Quick) { $Scales = @(1.0); $Sizes = @('1240x974', '1000x700'); $States = @('picked', 'multi', 'none', 'single-ok'); $Data = @('ref') }

# ---- the product, compiled exactly as the packer compiles it ---------------
. (Join-Path $Root 'build\sources.ps1')
$usings = New-Object System.Collections.Specialized.OrderedDictionary
$bodies = New-Object System.Text.StringBuilder
foreach ($f in $RdvSources) {
  $t = [IO.File]::ReadAllText((Join-Path $Root "src\csharp\$f"), [Text.Encoding]::UTF8)
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

# ---- the definition (read strictly, as the app reads it) and the ledger
# columns it names
$generic = -not [string]::IsNullOrEmpty($Settings)
$settingsPath = $(if ($generic) { (Resolve-Path -LiteralPath $Settings).Path } else { Join-Path $Root 'src\config\settings.json' })
$cfg = [Rdv3Config]::Load($settingsPath)
$screen = $cfg.Screen
$fields = New-Object Rdv3Fields (,$cfg.Data.ColumnRefs)
Write-Output ('definition: ' + $settingsPath)
Write-Output ('  ' + $cfg.Data.Describe())
Write-Output ('  ' + $screen.Describe())

# ---- what this definition calls its states and its judgment outcomes ---------
$workDef = $screen.Work
$storedInitial = $workDef.InitialStored
$storedLast = $storedInitial
$cur = $workDef.InitialState
for ($i = 0; $i -lt 8 -and $null -ne $cur; $i++) {
  $storedLast = $cur.Stored
  $tr = $workDef.FromState($cur.Id)
  $cur = $(if ($null -eq $tr) { $null } else { $workDef.ById($tr.To) })
}
$bandJudgment = $null
foreach ($sec in $screen.Sections) { if ($sec.Type -eq 'statusBand' -and $null -eq $bandJudgment) { $bandJudgment = $screen.JudgmentOf($sec.Judgment) } }
$judgeCol = $(if ($null -ne $bandJudgment -and $bandJudgment.Source.IsField) { $fields.IndexOf($bandJudgment.Source.Fields[0]) } else { -1 })
# a raw value that makes the band's judgment come out with the wanted look
function Judge-Look([string] $raw) {
  if ($judgeCol -lt 0) { return '' }
  $vw = New-Object Rdv3View
  $rec = New-Object 'string[]' $cfg.Data.Columns.Count
  for ($i = 0; $i -lt $rec.Length; $i++) { $rec[$i] = '' }
  $rec[$judgeCol] = $raw
  $vw.Record = $rec
  return ([Rdv3Eval]::Judge($bandJudgment, $vw, $fields)).Result.Look
}
function Value-For([string] $look, [string[]] $fallbacks) {
  if ($null -eq $bandJudgment) { return $fallbacks[0] }
  foreach ($r in $bandJudgment.Rules) {
    $cands = @()
    if ($r.EqualsAny.Length -gt 0) { $cands += $r.EqualsAny[0] }
    if ($r.Pattern.Length -gt 0) { $cands += ($r.Pattern -replace '[\^\$]', '') }
    foreach ($c in $cands) { if ((Judge-Look $c) -eq $look) { return $c } }
  }
  foreach ($c in $fallbacks) { if ((Judge-Look $c) -eq $look) { return $c } }
  throw ("no raw value of the band's judgment comes out as " + $look)
}
$valueOk = Value-For 'ok' @('OPEN', 'OK', 'DONE', '1', 'Y')
$valueNg = Value-For 'ng' @('HOLD', 'NG', '0', 'N')
$valueUndefined = Value-For 'undefined' @('WEIRD', 'ZZZ', '???', '')
Write-Output ('  judgment raw values: ok=' + $valueOk + ' ng=' + $valueNg + ' undefined=' + $valueUndefined + '   stored: ' + $storedInitial + ' -> ' + $storedLast)

# ---- sample data -----------------------------------------------------------
# a ledger line is 28 tab-separated columns: key1 key2 A[1..9] B[2..9] C[1..9]
function New-Line([string] $key1, [string] $key2, [hashtable] $o) {
  $f = New-Object 'string[]' 28
  for ($i = 0; $i -lt 28; $i++) { $f[$i] = '' }
  $f[0] = $key1; $f[1] = $key2
  $f[2] = $o.code; $f[3] = $o.name; $f[4] = $o.grade; $f[5] = $o.date; $f[6] = $o.amount
  $f[7] = $o.rate; $f[8] = $o.flag; $f[9] = $o.dept
  $f[11] = $o.slip; $f[12] = $o.bdate; $f[13] = $o.qty; $f[16] = $o.status
  $f[17] = $o.line; $f[18] = $o.memo; $f[19] = $o.item; $f[20] = $o.maker; $f[27] = $o.remark
  return ($f -join "`t")
}
# 'ref' repeats the reference's own sample strings; 'long' is deliberately
# longer than anything real
$refRow = @{ code = 'A52903'; name = 'CUSTOMER-0016168'; grade = 'B2'; date = '20240814';
  amount = '3969262'; rate = '0.1596'; flag = 'N'; dept = 'D568'; slip = 'SL00016168';
  bdate = '20240814'; qty = '320'; status = 'OPEN'; line = '001'; item = 'IT527901'; maker = 'MAKER-4128';
  memo = 'MEMO-85655'; remark = 'RMK-449812' }
$longRow = @{ code = 'A-4471-EXTENDED-CODE'; name = 'とても長い取引先名称株式会社 東日本第二営業統括本部 品質保証課'; grade = 'SPECIAL';
  date = '2024-11-08'; amount = '9,999,999,999'; rate = '0.8642'; flag = 'YES-LONG'; dept = '第二製造部 品質保証課 検査係';
  slip = 'SL-882140-EXT-0001'; bdate = '2026-08-14'; qty = '120,000'; status = 'HOLD'; line = '11482';
  item = 'IT-77120-LONG-ITEM'; maker = 'KYOWA-INDUSTRIAL-LONG';
  memo = ('長い摘要テキスト。' * 40); remark = ('長い備考テキスト。' * 40) }

# ---- generic sample rows, from the definition's own columns (-Settings) -----
# the first format the screen applies to a column says what shape its raw
# value has (a date in "from", a number); everything else is text
$fmtOf = @{}
foreach ($b in $screen.AllBindings()) {
  if ($b.IsField -and $null -ne $b.Format) { foreach ($f in $b.Fields) { if (-not $fmtOf.ContainsKey($f)) { $fmtOf[$f] = $b.Format } } }
}
function Generic-Value([string] $ref, [bool] $long) {
  $fmt = $fmtOf[$ref]
  if ($null -ne $fmt -and $fmt.Kind -eq 'date') { return (New-Object DateTime 2024, 8, 14, 16, 25, 0).ToString($fmt.From) }
  if ($null -ne $fmt -and $fmt.Kind -eq 'number') { return $(if ($long) { '9999999999' } else { '3969262' }) }
  $name = $ref.Substring($ref.IndexOf('.') + 1)
  if ($long) { return ('とても長い値 ' + $name + ' 東日本第二営業統括本部 品質保証課 検査係 ') * 3 }
  return $name.ToUpperInvariant() + '-0001'
}
function New-GenericLine([string] $key1, [string] $key2, [bool] $long, [string] $status) {
  $cols = $cfg.Data.Columns
  $f = New-Object 'string[]' $cols.Count
  for ($i = 0; $i -lt $cols.Count; $i++) { $f[$i] = Generic-Value $cols[$i].Ref $long }
  $f[$cfg.Data.IdentityCol] = $key2
  $f[$cfg.Data.SearchCol] = $key1
  if ($judgeCol -ge 0) {
    $f[$judgeCol] = $(switch ($status) { 'HOLD' { $valueNg } 'WEIRD' { $valueUndefined } default { $valueOk } })
  }
  return ($f -join "`t")
}

function New-Rows([int] $n, [hashtable] $o, [string] $key1, [string] $status) {
  $list = New-Object 'System.Collections.Generic.List[Rdv3CandRow]'
  for ($i = 0; $i -lt $n; $i++) {
    $r = New-Object Rdv3CandRow
    if ($generic) {
      $r.Line = New-GenericLine $key1 ('{0:D8}' -f (60919 + $i + 1)) ($o.ContainsKey('long')) $status
    } else {
      $h = $o.Clone(); if ($status) { $h.status = $status }
      $r.Line = New-Line $key1 ('{0:D8}' -f (60919 + $i + 1)) $h
    }
    $r.Stored = $(if ($i -eq 1) { $storedLast } else { $storedInitial })
    [void]$list.Add($r)
  }
  return $list
}
$longRow['long'] = $true

# ---- the form, never shown -------------------------------------------------
$form = New-Object Rdv3Form $screen
$null = $form.Handle
$form.Visible = $false
function Touch-Handles($c) { $null = $c.Handle; foreach ($k in $c.Controls) { Touch-Handles $k } }
Touch-Handles $form
$form.SetFields($fields)

function Set-Data([string] $kind) {
  if ($kind -eq 'long') {
    $form.SetIdentity('PID 182332', 'ReaderDataViewer-very-long-name.log')
    $form.SetLedger('ReaderDataViewer-Ledger-very-long-file-name.xlsx ・ 1,284,900 件', '1,284,900', '08-16 01:14')
    $form.SetWatch('とても長い名前の監視対象アプリケーション', '接続中（とても長いタイトルのテキストファイル - メモ帳）')
  } else {
    $form.SetIdentity('PID 18332', 'ReaderDataViewer.log')
    $form.SetLedger('ReaderDataViewer-Ledger.xlsx ・ 100,000 件', '100,000', '08-23 17:40')
    $form.SetWatch('メモ帳', '接続中（無題 - メモ帳）')
  }
}

function Set-State([string] $st, [string] $kind) {
  $o = if ($kind -eq 'long') { $longRow } else { $refRow }
  $key1 = if ($kind -eq 'long') { '40218764' } else { '00016168' }
  $form.ClearResult()
  Set-Data $kind
  $form.SetState('監視中')
  $form.EnableOps($true)
  switch ($st) {
    'idle'          { }
    'checking'      { $form.SetState('更新を確認中'); $form.EnableOps($false) }
    'single-ok'     { $form.SetKeyText($key1); $form.ShowCandidates($key1, (New-Rows 1 $o $key1 'OPEN'), 1); $form.SelectCandidate(0) }
    'single-ng'     { $form.SetKeyText($key1); $form.ShowCandidates($key1, (New-Rows 1 $o $key1 'HOLD'), 1); $form.SelectCandidate(0) }
    'multi'         { $form.SetKeyText($key1); $form.ShowCandidates($key1, (New-Rows 8 $o $key1 ''), 8) }
    'picked'        { $form.SetKeyText($key1); $form.ShowCandidates($key1, (New-Rows 8 $o $key1 'DONE'), 8); $form.SelectCandidate(2) }
    'saving'        { $form.SetKeyText($key1); $form.ShowCandidates($key1, (New-Rows 8 $o $key1 'DONE'), 8); $form.SelectCandidate(2)
                      $form.SetState('処理済を保存中'); $form.SetStoredState(2, $storedLast, $true) }
    'none'          { $form.SetKeyText($key1); $form.ShowCandidates($key1, (New-Rows 0 $o $key1 ''), 0) }
    'unknown-state' { $form.SetKeyText($key1); $form.ShowCandidates($key1, (New-Rows 1 $o $key1 'OPEN'), 1); $form.SelectCandidate(0)
                      $form.SetStoredState(0, 'MAYBE', $false) }
    'undefined'     { $form.SetKeyText($key1); $form.ShowCandidates($key1, (New-Rows 1 $o $key1 'WEIRD'), 1); $form.SelectCandidate(0) }
    'cleared'       { $form.SetKeyText('123'); $form.ClearResult() }
    default         { throw "unknown state: $st" }
  }
  [System.Windows.Forms.Application]::DoEvents()
}

# ---- the two checks --------------------------------------------------------
# containers legitimately hold other rectangles, so they are not paired
function Is-Container([string] $k) {
  if ($k -notmatch '\.') { return $true }                       # a section
  if ($k -match '^columns\d+\.\d+$') { return $true }           # a card inside columns
  if ($k -match '\.row\d+$') { return $true }                   # a field row
  if ($k -match '\.box$') { return $true }                      # a text box frame (the control sits in it)
  return $false
}
# an ellipsis is the design here, not a cut
function Is-Elastic([string] $k) {
  return ($k -match '\.row\d+\.value$') -or ($k -match '^statusBar\d+\.seg\d+$') -or ($k -match '\.title$') -or ($k -match '^statusBand\d+\.(sub|value)$')
}
# x/w of these is fixed by the layout; the rest follow their text
function Is-Structural([string] $k) {
  if ($k -match '^keyPanel\d+(\.label|\.value|\.input|\.btn\d+)$') { return $false }
  if ($k -match '^statusBand\d+\.(icon|value|sub)$') { return $false }
  if ($k -match '\.row\d+\.value$') { return $false }
  return $true
}

function Test-Integrity($el, [double] $cw, [double] $ch) {
  $keys = @($el.PSObject.Properties.Name | Where-Object { -not (Is-Container $_) })
  $n = $keys.Count
  $x0 = New-Object 'double[]' $n; $y0 = New-Object 'double[]' $n
  $x1 = New-Object 'double[]' $n; $y1 = New-Object 'double[]' $n
  for ($i = 0; $i -lt $n; $i++) {
    $r = $el.($keys[$i])
    $x0[$i] = $r[0]; $y0[$i] = $r[1]; $x1[$i] = $r[0] + $r[2]; $y1[$i] = $r[1] + $r[3]
  }
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

$refPath = Join-Path $Root 'docs\ui-ref-v2-geom.json'
$refGeom = (Get-Content -LiteralPath $refPath -Raw -Encoding UTF8 | ConvertFrom-Json).el

function Test-Fidelity($el) {
  $rows = New-Object System.Collections.ArrayList
  $bad = 0; $missing = 0
  foreach ($k in ($refGeom.PSObject.Properties.Name | Sort-Object)) {
    $r = $refGeom.$k
    $a = $el.$k
    if ($null -eq $a) { $missing++; continue }
    $d = @(($a[0] - $r[0]), ($a[1] - $r[1]), ($a[2] - $r[2]), ($a[3] - $r[3]))
    $worst = 0.0
    if ($k -match '\.row\d+\.value$') {
      # a right-aligned value: the reference measured the text span, the app
      # reports the box it is aligned in -- the right edge and the centre line
      # are what have to agree
      $cls = 'right'
      $dr = ($a[0] + $a[2]) - ($r[0] + $r[2])
      $dc = ($a[1] + $a[3] / 2) - ($r[1] + $r[3] / 2)
      $d = @($dr, $dc, 0, 0)
      $worst = [Math]::Max([Math]::Abs($dr), [Math]::Abs($dc))
    } else {
      if (Is-Structural $k) { $cls = 'geom'; $chk = @(0, 1, 2, 3) } else { $cls = 'text'; $chk = @(1, 3) }
      foreach ($c in $chk) { if ([Math]::Abs($d[$c]) -gt $worst) { $worst = [Math]::Abs($d[$c]) } }
    }
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
        # the body may be taller than the client and scroll: integrity is
        # judged on the laid-out card, not on the viewport
        $r = Test-Integrity $d.el $d.card[0] $d.card[1]
        $clip = @($d.clipped | Where-Object { -not (Is-Elastic $_) })
        $cases++
        if ($r.Overlap.Count -gt 0 -or $r.Outside.Count -gt 0 -or $clip.Count -gt 0) {
          $failed++
          Write-Output ('FAIL {0,-40} overlap {1} outside {2} clipped {3}' -f $tag, $r.Overlap.Count, $r.Outside.Count, $clip.Count)
          foreach ($h in $r.Overlap) { Write-Output ("       overlap  " + $h) }
          foreach ($h in $r.Outside) { Write-Output ("       outside  " + $h) }
          if ($clip.Count -gt 0) { Write-Output ("       clipped  " + ($clip -join ', ')) }
        }
        if ($st -eq 'single-ok' -and $kind -eq 'ref' -and $sz -eq '1240x974' -and $sc -eq 1.0) { $baseEl = $d.el }
      }
    }
  }
}
$form.Dispose()

Write-Output ''
Write-Output ('integrity : {0} cases, {1} failing   ({2:F0} s)' -f $cases, $failed, $sw.Elapsed.TotalSeconds)

$fidBad = 0
if ($generic) {
  Write-Output 'fidelity  : not measured for a definition other than the shipped one (integrity only)'
} elseif ($null -ne $baseEl) {
  $f = Test-Fidelity $baseEl
  $fidBad = $f.Bad + $f.Missing
  Write-Output ''
  Write-Output ("fidelity at the design size (tolerance {0} px)" -f $Tol)
  Write-Output ('{0,-24} {1,-5} {2,-28} {3,-28} {4}' -f 'element', 'class', 'reference x,y,w,h', 'app x,y,w,h', 'delta')
  foreach ($row in $f.Rows) {
    Write-Output ('{0,-24} {1,-5} {2,-28} {3,-28} {4}{5}' -f $row.Element, $row.Class, $row.Reference, $row.App, $row.Delta,
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
