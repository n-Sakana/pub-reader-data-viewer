# ============================================================================
# test_settings_geometry.ps1 -- the acceptance check for the settings modal.
#
#   powershell -File build\test_settings_geometry.ps1
#   powershell -File build\test_settings_geometry.ps1 -Quick
#   powershell -File build\test_settings_geometry.ps1 -Png     (also writes bitmaps)
#
# Same two conditions as the main screen (build\test_ui_geometry.ps1):
#
#   1. FIDELITY   at the design size and design scale, every named element sits
#                 within -Tol of the rectangle measured from the reference
#                 (docs\ui-ref-settings-geom.json, read out of
#                 "Reader Data Viewer_ver4.html" with the modal open).
#                 Layout decides y/h for every element and x/w for the
#                 structural ones; the rest follow their text, so a substituted
#                 font may legitimately move them.
#   2. INTEGRITY  at every scale, page and data set: no two leaf rectangles
#                 intersect, nothing falls outside the client area, and no
#                 string is silently cut off.
#
# The dialog is compiled from the SHIPPING sources and rendered WITHOUT EVER
# BEING SHOWN: the handle is created, DrawToBitmap paints it. No window
# appears, the foreground never moves, and what is measured is the product.
# ============================================================================
[CmdletBinding()]
param(
  [string] $Root = "",
  [string] $Out = "",
  [double[]] $Scales = @(1.0, 1.25, 1.5),
  [string[]] $Data = @('ref', 'long', 'empty', 'many'),
  [double] $Tol = 2.0,
  [switch] $Quick,
  [switch] $Png
)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
if ([string]::IsNullOrEmpty($Root)) { $Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path) }
if ([string]::IsNullOrEmpty($Out)) { $Out = Join-Path $Root 'work\ui-check-settings' }
if ($Quick) { $Scales = @(1.0); $Data = @('ref') }

# ---- the product, compiled exactly as the packer compiles it ---------------
$sources = @('Rdv3Core.cs', 'Rdv3Index.cs', 'Rdv3Ledger.cs', 'Rdv3Xlsx.cs', 'Rdv3Jobs.cs',
             'Rdv3Watch.cs', 'Rdv3Text.cs', 'Rdv3Geom.cs', 'Rdv3SetGeom.cs', 'Rdv3Json.cs',
             'Rdv3Config.cs', 'Rdv3Uia.cs', 'Rdv3Ui.cs', 'Rdv3Settings.cs', 'Rdv3App.cs')
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

# ---- the configurations the dialog is asked to show ------------------------
# 'ref' repeats the artifact's own sample, so the design-size comparison is
# like for like; the rest are the cases the reference never drew.
function New-Target([string] $name, [bool] $on, [string] $proc, [string] $cls,
                    [string] $like, [string] $fid, [string] $fcls, [string[]] $types) {
  $t = New-Object Rdv3Target
  $t.Name = $name
  $t.Enabled = $on
  $t.Window.ProcessName = $proc
  $t.Window.ClassName = $cls
  $t.Window.NameLike = $like
  $t.Field.AutomationId = $fid
  $t.Field.ClassName = $fcls
  $t.Field.ControlTypes = $types
  $t.Field.RequireValuePattern = $true
  return $t
}

function New-Config([string] $kind) {
  $c = [Rdv3Config]::Defaults()
  $c.Targets.Clear()
  if ($kind -eq 'long') {
    $c.SourcePath = 'C:\Program Files\Reader Data Viewer\very-long-folder-name\ReaderDataViewer.json'
    $t = New-Target 'とても長い監視対象の名前 第二製造部 検査端末' $true 'orderentrysystemclient' 'WindowsForms10.Window.8.app.0.1234567_r13_ad1' '*受発注システム - 伝票入力*' 'txtSlipNumberVeryLongIdentifier' 'WindowsForms10.EDIT.app.0.1234567' @('Edit', 'Document', 'Text')
    $s = New-Object Rdv3Match
    $s.AutomationId = 'pnlMainContainerWithAVeryLongName'
    $t.Steps.Add($s)
    $s2 = New-Object Rdv3Match
    $s2.AutomationId = 'tabDetailPage'
    $t.Steps.Add($s2)
    $c.Targets.Add($t)
    $c.DataDir = 'C:\Program Files\Reader Data Viewer\very-long-folder-name\data'
    $c.Ledger = 'ReaderDataViewer-Ledger-very-long-file-name.xlsx'
    $c.Log = 'ReaderDataViewer-very-long-name.log'
    $c.KeyLength = 12; $c.PollMs = 1000; $c.StableMs = 12000; $c.RebindMs = 60000
    $c.CandidateRowsShown = 999
  } elseif ($kind -eq 'empty') {
    $c.SourcePath = 'C:\Apps\ReaderDataViewer\ReaderDataViewer.json'
  } elseif ($kind -eq 'many') {
    $c.SourcePath = 'C:\Apps\ReaderDataViewer\ReaderDataViewer.json'
    for ($i = 1; $i -le 14; $i++) {
      $c.Targets.Add((New-Target ("監視対象 {0}" -f $i) ($i % 3 -ne 0) ('proc{0}' -f $i) 'Edit' '' ('fld{0}' -f $i) 'Edit' @('Edit')))
    }
  } else {
    $c.SourcePath = 'C:\Apps\ReaderDataViewer\ReaderDataViewer.json'
    $c.Targets.Add((New-Target 'メモ帳' $true 'notepad' 'Notepad' '*メモ帳' '15' 'Edit' @('Edit', 'Document')))
    $c.Targets.Add((New-Target '受発注システム' $true 'orderentry' '' '' 'txtSlipNo' '' @('Edit')))
    $c.Targets.Add((New-Target '旧検品端末' $false 'legacyscan' '' '' '' 'Edit' @('Edit')))
  }
  return $c
}

# ---- the two checks --------------------------------------------------------
# containers legitimately hold other rectangles, so they are not paired
$containers = @('head', 'head.rule', 'head.close', 'tabs', 'tabs.rule', 'tabs.t0', 'tabs.t1', 'tabs.t2',
                'body', 'foot', 'foot.rule', 'foot.save', 'p0.list', 'p0.btn.pick',
                'p0.read.seg', 'p0.scope.seg', 'p0.sec.window', 'p0.sec.field', 'p1.sec.key',
                'p1.sec.watch', 'p2.sec', 'p0.sec.window.rule', 'p0.sec.field.rule',
                'p1.sec.key.rule', 'p1.sec.watch.rule', 'p2.sec.rule', 'p0.step.rule')
for ($i = 0; $i -lt 20; $i++) { $containers += ('p0.card{0}' -f $i) }

# x/w of these is fixed by the layout; for everything else it follows the text
$structural = @('head', 'head.rule', 'head.close', 'tabs', 'tabs.rule', 'body', 'foot', 'foot.rule',
                'foot.save', 'foot.cancel',
                'p0.list', 'p0.btn.add', 'p0.btn.copy', 'p0.btn.del', 'p0.btn.pick', 'p0.hint',
                'p0.sec.window', 'p0.sec.window.rule', 'p0.sec.field', 'p0.sec.field.rule',
                'p0.win.id', 'p0.win.id.label', 'p0.win.class', 'p0.win.class.label',
                'p0.win.like', 'p0.win.like.label', 'p0.win.proc', 'p0.win.proc.label',
                'p0.fld.id', 'p0.fld.id.label', 'p0.fld.class', 'p0.fld.class.label',
                'p0.fld.types', 'p0.fld.types.label', 'p0.fld.index', 'p0.fld.index.label',
                'p0.read.seg', 'p0.read.seg0', 'p0.read.seg1', 'p0.read.seg2',
                'p0.scope.seg', 'p0.scope.seg0', 'p0.scope.seg1', 'p0.step.rule',
                'p1.sec.key', 'p1.sec.key.rule', 'p1.sec.watch', 'p1.sec.watch.rule',
                'p1.keylen', 'p1.poll', 'p1.stable', 'p1.rebind', 'p1.cand',
                'p1.key.note', 'p1.watch.note',
                'p2.sec', 'p2.sec.rule', 'p2.data', 'p2.data.label', 'p2.ledger', 'p2.ledger.label',
                'p2.log', 'p2.log.label', 'p2.note')

function Test-Integrity($el, [double] $cw, [double] $ch) {
  $keys = @($el.PSObject.Properties.Name | Where-Object { $containers -notcontains $_ })
  $n = $keys.Count
  $x0 = New-Object 'double[]' $n; $y0 = New-Object 'double[]' $n
  $x1 = New-Object 'double[]' $n; $y1 = New-Object 'double[]' $n
  for ($i = 0; $i -lt $n; $i++) {
    $r = $el.($keys[$i])
    $x0[$i] = $r[0]; $y0[$i] = $r[1]; $x1[$i] = $r[0] + $r[2]; $y1[$i] = $r[1] + $r[3]
  }
  $order = 0..([Math]::Max(1, $n) - 1) | Sort-Object { $y0[$_] }
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
  # ...and a part must stay inside the thing it is part of. Overlap alone does
  # not catch a card's "enabled" tag pushed past the card's own right edge by a
  # long name: it lands in the list box, which is nobody's neighbour.
  $names = @($el.PSObject.Properties.Name)
  foreach ($k in $names) {
    $parent = ''
    foreach ($p in $names) {
      if ($p -ne $k -and $k.StartsWith($p) -and ($containers -contains $p) -and $p.Length -gt $parent.Length) {
        $parent = $p
      }
    }
    if ($parent -eq '') { continue }
    $r = $el.$k; $q = $el.$parent
    if ($r[0] -lt ($q[0] - 1.0) -or $r[1] -lt ($q[1] - 1.0) -or
        ($r[0] + $r[2]) -gt ($q[0] + $q[2] + 1.0) -or ($r[1] + $r[3]) -gt ($q[1] + $q[3] + 1.0)) {
      [void]$outside.Add(('{0} escapes {1} [{2},{3},{4},{5}] vs [{6},{7},{8},{9}]' -f
        $k, $parent, $r[0], $r[1], $r[2], $r[3], $q[0], $q[1], $q[2], $q[3]))
    }
  }
  return @{ Overlap = $hits; Outside = $outside }
}

$refDoc = Get-Content -LiteralPath (Join-Path $Root 'docs\ui-ref-settings-geom.json') -Raw -Encoding UTF8 | ConvertFrom-Json

function Test-Fidelity($el, $refGeom, [string] $prefix) {
  $rows = New-Object System.Collections.ArrayList
  $bad = 0; $missing = 0
  foreach ($k in ($refGeom.PSObject.Properties.Name | Sort-Object)) {
    # chrome keys belong to every page; a "pN." key only to its own
    if ($prefix -ne '' -and ($k -match '^p[0-9]\.') -and -not $k.StartsWith($prefix)) { continue }
    # the selected-tab underline moves with the page; the reference measured it
    # under the first tab, so it is only comparable there
    if ($k -eq 'tabs.mark' -and $prefix -ne '' -and $prefix -ne 'p0.') { continue }
    $r = $refGeom.$k
    $a = $el.$k
    if ($null -eq $a) { $missing++
      [void]$rows.Add([pscustomobject]@{ Element = $k; Class = 'MISSING'
        Reference = ('{0},{1},{2},{3}' -f $r[0], $r[1], $r[2], $r[3]); App = '-'; Delta = '-'
        Worst = 999; Over = $true })
      continue }
    $d = @(($a[0] - $r[0]), ($a[1] - $r[1]), ($a[2] - $r[2]), ($a[3] - $r[3]))
    if ($structural -contains $k) { $cls = 'geom'; $chk = @(0, 1, 2, 3) }
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

function Touch-Handles($c) {
  $null = $c.Handle
  foreach ($k in $c.Controls) { Touch-Handles $k }
}

# ---- run -------------------------------------------------------------------
$cases = 0; $failed = 0
$basePage = @{}
$sw = [Diagnostics.Stopwatch]::StartNew()
foreach ($kind in $Data) {
  $cfg = New-Config $kind
  foreach ($sc in $Scales) {
    # ShowDialog is never called: the constructor builds the dialog, the handle
    # is created here, and DrawToBitmap paints it into a bitmap
    $dlg = [Rdv3SettingsForm]::ForCheck($cfg)
    $null = $dlg.Handle
    $dlg.Visible = $false
    Touch-Handles $dlg
    $dlg.SetDesignScale([single]$sc)
    for ($p = 0; $p -lt 3; $p++) {
      $dlg.GoToPage($p)
      [System.Windows.Forms.Application]::DoEvents()
      $tag = 'p{0}-{1}-{2}' -f $p, $kind, ($sc.ToString('0.00'))
      $bmp = New-Object Drawing.Bitmap $dlg.Width, $dlg.Height
      $dlg.DrawToBitmap($bmp, (New-Object Drawing.Rectangle 0, 0, $dlg.Width, $dlg.Height))
      $json = $dlg.GeometryDump()
      [IO.File]::WriteAllText((Join-Path $Out ("geom-$tag.json")), $json, (New-Object Text.UTF8Encoding $false))
      if ($Png) {
        $o = $dlg.PointToScreen([Drawing.Point]::Empty)
        $crop = New-Object Drawing.Rectangle ($o.X - $dlg.Left), ($o.Y - $dlg.Top), $dlg.ClientSize.Width, $dlg.ClientSize.Height
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
        Write-Output ('FAIL {0,-24} overlap {1} outside {2} clipped {3}' -f $tag, $r.Overlap.Count, $r.Outside.Count, $clip.Count)
        foreach ($h in $r.Overlap) { Write-Output ("       overlap  " + $h) }
        foreach ($h in $r.Outside) { Write-Output ("       outside  " + $h) }
        if ($clip.Count -gt 0) { Write-Output ("       clipped  " + ($clip -join ', ')) }
      }
      if ($kind -eq 'ref' -and $sc -eq 1.0) { $basePage[$p] = $d.el }
    }
    $dlg.Dispose()
  }
}

# ---- the picker ------------------------------------------------------------
$pkCases = 0; $pkFailed = 0
foreach ($sc in $Scales) {
  $pk = [Rdv3PickerForm]::ForCheck()
  $null = $pk.Handle
  $pk.Visible = $false
  Touch-Handles $pk
  $pk.SetDesignScale([single]$sc)
  foreach ($kind in @('ref', 'long', 'blank')) {
    if ($kind -eq 'long') {
      $pk.SetSample('DataItem', 'txtSlipNumberVeryLongAutomationIdentifier',
        'WindowsForms10.EDIT.app.0.1234567_r13_ad1', '行 3 「SL00016168 とても長い項目名称の見出し」',
        'orderentrysystemclient', '')
    } elseif ($kind -eq 'blank') {
      $pk.SetSample('', '', '', '', '', '')
    } else {
      $pk.SetSample('DataItem', '', 'SysListView32', '行 3 「SL00016168」', 'orderentry', '')
    }
    [System.Windows.Forms.Application]::DoEvents()
    $tag = 'pk-{0}-{1}' -f $kind, ($sc.ToString('0.00'))
    $bmp = New-Object Drawing.Bitmap $pk.Width, $pk.Height
    $pk.DrawToBitmap($bmp, (New-Object Drawing.Rectangle 0, 0, $pk.Width, $pk.Height))
    $json = $pk.GeometryDump()
    [IO.File]::WriteAllText((Join-Path $Out ("geom-$tag.json")), $json, (New-Object Text.UTF8Encoding $false))
    if ($Png) {
      $o = $pk.PointToScreen([Drawing.Point]::Empty)
      $crop = New-Object Drawing.Rectangle ($o.X - $pk.Left), ($o.Y - $pk.Top), $pk.ClientSize.Width, $pk.ClientSize.Height
      $c2 = $bmp.Clone($crop, $bmp.PixelFormat)
      $c2.Save((Join-Path $Out ("shot-$tag.png")), [Drawing.Imaging.ImageFormat]::Png)
      $c2.Dispose()
    }
    $bmp.Dispose()
    $d = $json | ConvertFrom-Json
    $r = Test-Integrity $d.el $d.client[0] $d.client[1]
    $clip = @($d.clipped)
    $pkCases++
    if ($r.Overlap.Count -gt 0 -or $r.Outside.Count -gt 0 -or $clip.Count -gt 0) {
      $pkFailed++
      Write-Output ('FAIL {0,-24} overlap {1} outside {2} clipped {3}' -f $tag, $r.Overlap.Count, $r.Outside.Count, $clip.Count)
      foreach ($h in $r.Overlap) { Write-Output ("       overlap  " + $h) }
      foreach ($h in $r.Outside) { Write-Output ("       outside  " + $h) }
      if ($clip.Count -gt 0) { Write-Output ("       clipped  " + ($clip -join ', ')) }
    }
    if ($kind -eq 'ref' -and $sc -eq 1.0) { $script:pkBase = $d.el }
  }
  $pk.Dispose()
}

Write-Output ''
Write-Output ('integrity : {0} modal + {1} picker cases, {2} failing   ({3:F0} s)' -f
  $cases, $pkCases, ($failed + $pkFailed), $sw.Elapsed.TotalSeconds)

# ---- fidelity at the design size ------------------------------------------
$fidBad = 0
$allRows = New-Object System.Collections.ArrayList
for ($p = 0; $p -lt 3; $p++) {
  if ($null -eq $basePage[$p]) { continue }
  $f = Test-Fidelity $basePage[$p] $refDoc.el ('p{0}.' -f $p)
  $fidBad += $f.Bad + $f.Missing
  foreach ($row in $f.Rows) { [void]$allRows.Add($row) }
}
if ($null -ne $script:pkBase) {
  $f = Test-Fidelity $script:pkBase $refDoc.pk ''
  $fidBad += $f.Bad + $f.Missing
  foreach ($row in $f.Rows) { [void]$allRows.Add($row) }
}

Write-Output ''
Write-Output ("fidelity at the design size (tolerance {0} px)" -f $Tol)
Write-Output ('{0,-24} {1,-8} {2,-26} {3,-26} {4}' -f 'element', 'class', 'reference x,y,w,h', 'app x,y,w,h', 'delta')
foreach ($row in $allRows) {
  Write-Output ('{0,-24} {1,-8} {2,-26} {3,-26} {4}{5}' -f $row.Element, $row.Class, $row.Reference, $row.App, $row.Delta,
    $(if ($row.Over) { '   <-- over' } else { '' }))
}
Write-Output ('elements {0}, outside tolerance {1}' -f $allRows.Count, $fidBad)

Write-Output ''
if (($failed + $pkFailed) -eq 0 -and $fidBad -eq 0) { Write-Output 'RESULT: PASS'; exit 0 }
Write-Output ('RESULT: FAIL (integrity {0}, fidelity {1})' -f ($failed + $pkFailed), $fidBad)
exit 1
