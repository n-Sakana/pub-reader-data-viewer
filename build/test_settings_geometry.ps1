# ============================================================================
# test_settings_geometry.ps1 -- the acceptance check for the overlays: the
# settings modal, the candidate list modal, the confirm modal and the picker.
#
#   powershell -File build\test_settings_geometry.ps1
#   powershell -File build\test_settings_geometry.ps1 -Png     (also writes bitmaps)
#
# Each window is built from the shipping sources and rendered WITHOUT EVER
# BEING SHOWN (DrawToBitmap). Two conditions:
#
#   1. FIDELITY   at the design scale, the named elements of the settings modal
#                 and the picker sit within -Tol of the rectangles measured from
#                 the v2 reference (the table below, in CSS px of the modal).
#   2. INTEGRITY  at three scaling factors and several data sets: no two leaf
#                 rectangles intersect, nothing falls outside the window, no
#                 string is silently cut (card names and summaries may
#                 ellipsise by design).
# ============================================================================
[CmdletBinding()]
param(
  [string] $Root = "",
  [string] $Out = "",
  [double[]] $Scales = @(1.0, 1.25, 1.5),
  [double] $Tol = 2.0,
  [switch] $Png
)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
if ([string]::IsNullOrEmpty($Root)) { $Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path) }
if ([string]::IsNullOrEmpty($Out)) { $Out = Join-Path $Root 'work\ui-check' }

. (Join-Path $Root 'build\sources.ps1')
$usings = New-Object System.Collections.Specialized.OrderedDictionary
$bodies = New-Object System.Text.StringBuilder
foreach ($f in $RdvSources) {
  $t = [IO.File]::ReadAllText((Join-Path $Root "src\csharp\$f"), [Text.Encoding]::UTF8)
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
try { Add-Type -TypeDefinition $cs -ReferencedAssemblies $refs -Language CSharp }
catch { Write-Output 'COMPILE FAILED'; Write-Output $_.Exception.Message; exit 1 }
Write-Output 'compile ok'
[void][System.IO.Directory]::CreateDirectory($Out)

# ---- the reference (docs\ui-reference\v2.html, measured in Chromium) -------
# settings modal, CSS px relative to the modal (560 x 769.5)
$refSettings = @{
  'st.title'        = @(20, 16, 40, 31)
  'st.close'        = @(510, 16.5, 30, 30)
  'st.sec.targets'  = @(20, 63, 520, 20.1)
  'st.cards'        = @(20, 91.1, 520, 64)
  'st.pick'         = @(20, 165.1, 520, 38)
  'st.note.target'  = @(20, 211.1, 520, 21.7)
  'st.rule1'        = @(20, 250.8, 520, 1)
  'st.sec.pattern'  = @(20, 267.8, 520, 20.1)
  'st.pattern'      = @(20, 298, 520, 40)
  'st.note.pattern' = @(20, 346, 520, 20.1)
  'st.rule2'        = @(20, 384.1, 520, 1)
  'st.sec.files'    = @(20, 401.1, 520, 20.1)
  'st.data.label'   = @(20, 431.2, 520, 21.7)
  'st.data'         = @(20, 458.9, 520, 40)
  'st.ledger.label' = @(20, 510.9, 520, 21.7)
  'st.ledger'       = @(20, 538.6, 520, 40)
  'st.log.label'    = @(20, 590.6, 520, 21.7)
  'st.log'          = @(20, 618.3, 520, 40)
  'st.note.files'   = @(20, 666.3, 520, 20.1)
  'st.foot'         = @(0, 706.5, 560, 63)
  'st.cancel'       = @(396, 719.5, 62, 38)
  'st.save'         = @(466, 719.5, 74, 38)
}
# the picker, relative to its panel (460 x 323.5)
$refPicker = @{
  'pk.title' = @(14.3, 14.3, 119.1, 34.1)
  'pk.how'   = @(14.3, 56.4, 431.4, 23.3)
  'pk.rule'  = @(14.3, 89.7, 431.4, 1)
  'pk.k0'    = @(14.3, 104.6, 120, 21.7)
  'pk.v0'    = @(134.3, 103.9, 311.4, 23.3)
  'pk.k5'    = @(14.3, 242.6, 120, 21.7)
  'pk.v5'    = @(134.3, 241.9, 311.4, 23.3)
  'pk.esc'   = @(14.3, 283.4, 62.6, 21.7)
  'pk.close' = @(373.7, 279.3, 72, 30)
}
# structural on x/y/w/h; the rest (buttons, the note, the title) follow their text
$textish = @('st.title', 'st.cancel', 'st.save', 'st.note.target', 'pk.title', 'pk.esc', 'pk.close')

function Touch-Handles($c) { $null = $c.Handle; foreach ($k in $c.Controls) { Touch-Handles $k } }

function Render($form, [string] $name) {
  $null = $form.Handle
  Touch-Handles $form
  [System.Windows.Forms.Application]::DoEvents()
  $bmp = New-Object Drawing.Bitmap $form.Width, $form.Height
  $form.DrawToBitmap($bmp, (New-Object Drawing.Rectangle 0, 0, $form.Width, $form.Height))
  if ($Png) { $bmp.Save((Join-Path $Out ("modal-$name.png")), [Drawing.Imaging.ImageFormat]::Png) }
  $bmp.Dispose()
  $json = $form.GeometryDump()
  [IO.File]::WriteAllText((Join-Path $Out ("modal-$name.json")), $json, (New-Object Text.UTF8Encoding $false))
  return ($json | ConvertFrom-Json)
}

function Is-Container([string] $k) {
  return ($k -eq 'st.head') -or ($k -eq 'st.foot') -or ($k -eq 'cl.head') -or ($k -eq 'cl.table') -or ($k -eq 'st.cards')
}
function Is-Elastic([string] $k) { return ($k -match '^st\.note\.') -or ($k -match '^cl\.hint$') }

function Test-Integrity($el, [double] $cw, [double] $ch) {
  $keys = @($el.PSObject.Properties.Name | Where-Object { -not (Is-Container $_) })
  $hits = New-Object System.Collections.ArrayList
  for ($a = 0; $a -lt $keys.Count; $a++) {
    $p = $el.($keys[$a])
    for ($b = $a + 1; $b -lt $keys.Count; $b++) {
      $q = $el.($keys[$b])
      $w = [Math]::Min($p[0] + $p[2], $q[0] + $q[2]) - [Math]::Max($p[0], $q[0])
      $h = [Math]::Min($p[1] + $p[3], $q[1] + $q[3]) - [Math]::Max($p[1], $q[1])
      if ($w -gt 0.6 -and $h -gt 0.6) { [void]$hits.Add(('{0} x {1}' -f $keys[$a], $keys[$b])) }
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

function Test-Fidelity($el, $ref, [string] $what) {
  $bad = 0
  foreach ($k in ($ref.Keys | Sort-Object)) {
    $r = $ref[$k]; $a = $el.$k
    if ($null -eq $a) { $line = ("  {0,-18} MISSING" -f $k); [Console]::Out.WriteLine($line); $bad++; continue }
    $chk = if ($textish -contains $k) { @(1, 3) } else { @(0, 1, 2, 3) }
    $worst = 0.0
    foreach ($c in $chk) { $d = [Math]::Abs($a[$c] - $r[$c]); if ($d -gt $worst) { $worst = $d } }
    $flag = if ($worst -gt $Tol) { '   <-- over' } else { '' }
    if ($worst -gt $Tol) { $bad++ }
    $line = ('  {0,-18} ref {1,-24} app {2,-24}{3}' -f $k, ($r -join ','), ($a -join ','), $flag)
    [Console]::Out.WriteLine($line)
  }
  $line = ("{0}: {1} elements, {2} outside tolerance" -f $what, $ref.Count, $bad)
  [Console]::Out.WriteLine($line)
  return $bad
}

# ---- data --------------------------------------------------------------------
$cfg0 = [Rdv3Config]::Load((Join-Path $Root 'src\config\settings.json'))
$screen = $cfg0.Screen
$fields = New-Object Rdv3Fields (,$cfg0.Data.ColumnRefs)

function New-Cfg([int] $targets, [bool] $long) {
  $c = $cfg0.Clone()
  $c.Targets.Clear()
  for ($i = 0; $i -lt $targets; $i++) {
    $t = New-Object Rdv3Target
    $t.Name = $(if ($long) { ('とても長い名前の監視対象アプリケーション番号 {0} 在庫照会画面' -f ($i + 1)) } else { ('対象 {0}' -f ($i + 1)) })
    $t.Window.ProcessName = $(if ($long) { 'VeryLongProcessNameOfTheApplication' } else { 'LobApp' })
    $t.Field.AutomationId = $(if ($long) { 'txtBarcodeWithAVeryLongAutomationIdentifier' } else { 'txtBarcode' })
    $c.Targets.Add($t)
  }
  if ($long) { $c.DataDir = 'D:\とても長いフォルダー名\data\更に長いフォルダー名\csv'; $c.KeyPattern = '^[0-9A-Z]{8,16}$' }
  return $c
}

function New-Rows([int] $n, [bool] $long) {
  $list = New-Object 'System.Collections.Generic.List[Rdv3CandRow]'
  $st = @('OPEN', 'DONE', 'HOLD', 'VOID')
  for ($i = 0; $i -lt $n; $i++) {
    $f = New-Object 'string[]' 28
    for ($k = 0; $k -lt 28; $k++) { $f[$k] = '' }
    $f[0] = '00001001'; $f[1] = ('{0:D8}' -f (30179 + $i * 911))
    $f[11] = $(if ($long) { 'SL-882140-EXT-0001-LONG' } else { 'SL00001001' })
    $f[12] = '20240326'; $f[13] = $(if ($long) { '1,200,000' } else { '427' }); $f[16] = $st[$i % 4]; $f[17] = ('{0:D3}' -f ($i + 1))
    $f[19] = $(if ($long) { 'IT-77120-LONG-ITEM-CODE' } else { 'IT966839' }); $f[20] = $(if ($long) { 'KYOWA-INDUSTRIAL-LONG-MAKER' } else { 'MAKER-9815' })
    $r = New-Object Rdv3CandRow
    $r.Line = ($f -join "`t")
    $r.Stored = $(if ($i -eq 1) { 'TRUE' } elseif ($i -eq 2) { 'MAYBE' } else { 'FALSE' })
    [void]$list.Add($r)
  }
  return $list
}

# ---- run -----------------------------------------------------------------------
$cases = 0; $failed = 0; $fidBad = 0
foreach ($sc in $Scales) {
  foreach ($case in @(@{ n = 1; long = $false }, @{ n = 3; long = $true }, @{ n = 0; long = $false }, @{ n = 6; long = $false })) {
    $cfg = New-Cfg $case.n $case.long
    $st = [Rdv3SettingsForm]::ForCheck($cfg, [single]$sc)
    $tag = ('settings-{0}-{1}-{2:0.00}' -f $case.n, $(if ($case.long) { 'long' } else { 'ref' }), $sc)
    $d = Render $st $tag
    $r = Test-Integrity $d.el $d.client[0] $d.client[1]
    $clip = @($d.clipped | Where-Object { -not (Is-Elastic $_) })
    $cases++
    if ($r.Overlap.Count -gt 0 -or $r.Outside.Count -gt 0 -or $clip.Count -gt 0) {
      $failed++
      Write-Output ('FAIL {0,-34} overlap {1} outside {2} clipped {3}' -f $tag, $r.Overlap.Count, $r.Outside.Count, $clip.Count)
      foreach ($h in $r.Overlap) { Write-Output ("       overlap  " + $h) }
      foreach ($h in $r.Outside) { Write-Output ("       outside  " + $h) }
      if ($clip.Count -gt 0) { Write-Output ("       clipped  " + ($clip -join ', ')) }
    }
    if ($sc -eq 1.0 -and $case.n -eq 1 -and -not $case.long) {
      Write-Output ''
      Write-Output 'settings modal at the design scale:'
      $fidBad += Test-Fidelity $d.el $refSettings 'settings'
    }
    $st.Dispose()
  }

  foreach ($case in @(@{ n = 1; long = $false }, @{ n = 6; long = $false }, @{ n = 14; long = $true })) {
    $rows = New-Rows $case.n $case.long
    $cl = [Rdv3CandidatesForm]::ForCheck($screen.Candidates, $fields, $screen.Work, $rows, $case.n, 2, [single]$sc)
    $tag = ('candidates-{0}-{1}-{2:0.00}' -f $case.n, $(if ($case.long) { 'long' } else { 'ref' }), $sc)
    $d = Render $cl $tag
    $r = Test-Integrity $d.el $d.client[0] $d.client[1]
    $clip = @($d.clipped | Where-Object { -not (Is-Elastic $_) })
    $cases++
    if ($r.Overlap.Count -gt 0 -or $r.Outside.Count -gt 0 -or $clip.Count -gt 0) {
      $failed++
      Write-Output ('FAIL {0,-34} overlap {1} outside {2} clipped {3}' -f $tag, $r.Overlap.Count, $r.Outside.Count, $clip.Count)
      foreach ($h in $r.Overlap) { Write-Output ("       overlap  " + $h) }
      foreach ($h in $r.Outside) { Write-Output ("       outside  " + $h) }
      if ($clip.Count -gt 0) { Write-Output ("       clipped  " + ($clip -join ', ')) }
    }
    $cl.Dispose()
  }

  foreach ($body in @('表示中のレコード (番号2 = 00097542) を処理済にします。よろしいですか?',
                      ("CSV に変更があります。統合台帳を更新しますか?`n(作業状態は変更のないレコードへ引き継がれます)"),
                      ('保存済みの統合台帳が読めません:' + "`n" + 'ledger header row does not match the state column and the 28 CSV columns: C:\very\long\path\ReaderDataViewer-Ledger.xlsx' + "`n" + 'CSV から作り直しますか? (作業状態は失われます)'))) {
    $cf = [Rdv3ConfirmForm]::ForCheck('更新の確認', $body, [single]$sc)
    $tag = ('confirm-{0}-{1:0.00}' -f $body.Length, $sc)
    $d = Render $cf $tag
    $r = Test-Integrity $d.el $d.client[0] $d.client[1]
    $cases++
    if ($r.Overlap.Count -gt 0 -or $r.Outside.Count -gt 0 -or $d.clipped.Count -gt 0) {
      $failed++
      Write-Output ('FAIL {0,-34} overlap {1} outside {2} clipped {3}' -f $tag, $r.Overlap.Count, $r.Outside.Count, $d.clipped.Count)
      foreach ($h in $r.Overlap) { Write-Output ("       overlap  " + $h) }
      foreach ($h in $r.Outside) { Write-Output ("       outside  " + $h) }
    }
    $cf.Dispose()
  }

  foreach ($case in @('ref', 'long', 'none')) {
    $pk = [Rdv3PickerForm]::ForCheck([single]$sc)
    if ($case -eq 'ref') { $pk.SetSample('Document', '15', 'RichEditD2DPT', 'テキスト エディター', 'Notepad', '00016168') }
    elseif ($case -eq 'long') { $pk.SetSample('DataItem', 'dgvOrderLines_Row12_ColBarcodeValueCell', 'WindowsForms10.Window.8.app.0.2bf8098_r9_ad1', 'とても長い名前のデータグリッドのセル 受注明細 12 行目 バーコード列', 'VeryLongProcessName', '0000016168000001616800000161680000016168') }
    else { $pk.SetSample('', '', '', '', '', '') }
    $tag = ('picker-{0}-{1:0.00}' -f $case, $sc)
    $d = Render $pk $tag
    $r = Test-Integrity $d.el $d.client[0] $d.client[1]
    $clip = @($d.clipped | Where-Object { $_ -notmatch '^pk\.v\d$' })
    $cases++
    if ($r.Overlap.Count -gt 0 -or $r.Outside.Count -gt 0 -or $clip.Count -gt 0) {
      $failed++
      Write-Output ('FAIL {0,-34} overlap {1} outside {2} clipped {3}' -f $tag, $r.Overlap.Count, $r.Outside.Count, $clip.Count)
      foreach ($h in $r.Overlap) { Write-Output ("       overlap  " + $h) }
      foreach ($h in $r.Outside) { Write-Output ("       outside  " + $h) }
      if ($clip.Count -gt 0) { Write-Output ("       clipped  " + ($clip -join ', ')) }
    }
    if ($sc -eq 1.0 -and $case -eq 'ref') {
      Write-Output ''
      Write-Output 'picker at the design scale:'
      $fidBad += Test-Fidelity $d.el $refPicker 'picker'
    }
    $pk.Dispose()
  }
}

Write-Output ''
Write-Output ('integrity : {0} cases, {1} failing' -f $cases, $failed)
Write-Output ('fidelity  : {0} outside tolerance' -f $fidBad)
if ($failed -eq 0 -and $fidBad -eq 0) { Write-Output 'RESULT: PASS'; exit 0 }
Write-Output 'RESULT: FAIL'
exit 1
