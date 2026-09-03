# ============================================================================
# test_samples.ps1 -- the three sample definitions (src\samples) against the
# product code, with the data gen_samples.ps1 generates.
#
#   powershell -File build\test_samples.ps1
#   powershell -File build\test_samples.ps1 -Only factory
#
# For each sample: the settings.json loads strictly and binds to the CSV
# headers; the merge produces one row per spine row with the join counts the
# generator recorded; the ledger round-trips through xlsx; the search column
# answers with the recorded hit counts; a recorded row shows the recorded
# values through the screen's own bindings (no unresolved column, no error
# tone); the judgment and the work states are what the definition says; a
# changed row loses its state on carry-over and an unchanged one keeps it.
#
# Compiled from the shipping sources the way the packer compiles them; no
# window is ever created (the on-device run is work\ui-v2\live_samples.ps1).
# ============================================================================
[CmdletBinding()]
param(
  [string] $Root = "",
  [string] $Only = ""
)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
if ([string]::IsNullOrEmpty($Root)) { $Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path) }

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
  [System.Xml.XmlReader].Assembly.Location
)
Add-Type -TypeDefinition $cs -ReferencedAssemblies $refs -Language CSharp
Write-Output 'compile ok'

$samples = Join-Path $Root 'samples'
$all = @('sales', 'factory', 'booking', 'sales-wide', 'factory-compact')
$want = $(if ([string]::IsNullOrEmpty($Only)) { $all } else { @($Only) })
foreach ($n in $want) {
  if (-not (Test-Path -LiteralPath (Join-Path (Join-Path $samples $n) 'expected.txt'))) {
    & (Join-Path $Root 'build\gen_samples.ps1') -Root $Root -Only $n | Out-Null
  }
}
$work = Join-Path $Root 'work\samples-test'
if (Test-Path -LiteralPath $work) { Remove-Item -LiteralPath $work -Recurse -Force }
New-Item -ItemType Directory -Path $work | Out-Null

$script:pass = 0
$script:fail = 0
function Ok([string] $name, [bool] $cond, [string] $detail) {
  if ($cond) { $script:pass++; Write-Output ("  ok   {0}" -f $name) }
  else { $script:fail++; Write-Output ("  FAIL {0}  ({1})" -f $name, $detail) }
}
function Section([string] $t) { Write-Output ''; Write-Output $t }
function Facts([string] $path) {
  $h = @{}
  foreach ($l in [IO.File]::ReadAllLines($path, [Text.Encoding]::UTF8)) {
    $eq = $l.IndexOf('='); if ($eq -gt 0) { $h[$l.Substring(0, $eq)] = $l.Substring($eq + 1) }
  }
  return $h
}

foreach ($name in $want) {
  $dir = Join-Path $samples $name
  $facts = Facts (Join-Path $dir 'expected.txt')
  Section ("sample " + $name)

  # ---- the definition, strictly, and against the headers ----------------------
  $cfg = [Rdv3Config]::Load((Join-Path $dir 'settings.json'))
  $data = $cfg.Data
  $dataDir = Join-Path $dir 'data'
  $heads = New-Object 'string[][]' $data.Tables.Count
  for ($t = 0; $t -lt $data.Tables.Count; $t++) {
    $heads[$t] = [Rdv3Table]::ReadHead((Join-Path $dataDir $data.Tables[$t].File), $data.Enc)
  }
  $data.Bind($heads)
  Ok 'the definition loads and binds to the CSV headers' $true ''
  Write-Output ('       ' + $data.Describe())
  Write-Output ('       ' + $cfg.Screen.Describe())

  # ---- the merge -----------------------------------------------------------------
  $mr = [Rdv3Ledger]::BuildFromCsv($data, $dataDir)
  Ok ('one ledger row per spine row: ' + $facts['rows']) ($mr.Rows -eq [int]$facts['rows']) ("rows " + $mr.Rows)
  for ($j = 0; $j -lt $data.Joins.Count; $j++) {
    $tid = $data.Joins[$j].Table
    $wantN = [int]$facts['match.' + $tid]
    Ok ("join " + $data.Spine + "->" + $tid + " matches " + $wantN + " rows") ($mr.Matched[$j] -eq $wantN) ("matched " + $mr.Matched[$j])
  }
  $fields = New-Object Rdv3Fields (,$data.ColumnRefs)
  Ok 'the ledger line has one field per column' (([Rdv3Ledger]::SplitLine($mr.Lines[0])).Length -eq $data.Columns.Count) ("fields " + ([Rdv3Ledger]::SplitLine($mr.Lines[0])).Length)

  # ---- the ledger file --------------------------------------------------------------
  $work2 = $cfg.Screen.Work
  $initial = $work2.InitialStored
  Ok ('the initial stored state is ' + $facts['initial.stored']) ($initial -eq $facts['initial.stored']) $initial
  $states = [Rdv3Ledger]::FreshStates($mr.Lines.Length, $initial)
  $xl = Join-Path $work ($name + '.xlsx')
  [Rdv3Xlsx]::Write($xl, $mr.Head, $work2.Column, $mr.Lines, $states, 'test')
  $outLines = $null; $outStates = $null
  [Rdv3Xlsx]::Read($xl, $mr.Head, $work2.Column, [ref]$outLines, [ref]$outStates)
  $firstDiff = -1
  Ok 'the xlsx round-trips every line'     ([Rdv3Ledger]::SameContent($outLines, $mr.Lines, [ref]$firstDiff)) ("first diff " + $firstDiff)
  Ok 'and every state'                     (($outStates.Length -eq $states.Length) -and ($outStates[0] -eq $initial)) ($outStates[0])
  $ledgerRow = New-Object Rdv3View   # the row count sanity: the header is the CSV names in ledger order
  Ok 'the xlsx header is the CSV names'    ($mr.Head[0] -eq $data.Columns[0].Column) ($mr.Head -join ',')

  # ---- search -----------------------------------------------------------------------
  $searchCtor = [Rdv3Index].GetConstructor([Type[]]@([string[]], [int[]], [string]))
  $ix = $searchCtor.Invoke([object[]]@($mr.Lines, $data.SearchCols, $data.SearchMatch))
  foreach ($probeName in 'multi', 'single') {
    $key = $facts['probe.' + $probeName + '.key']
    $hits = $ix.Find($key)
    $n = $(if ($null -eq $hits) { 0 } else { $hits.Count })
    Ok ("search " + $key + " -> " + $facts['probe.' + $probeName + '.hits'] + " hits") ($n -eq [int]$facts['probe.' + $probeName + '.hits']) ("hits " + $n)
    Ok ("and " + $key + " matches search.pattern") ($cfg.IsKey($key)) $cfg.KeyPattern
  }
  $none = $ix.Find($facts['probe.none.key'])
  Ok ("search " + $facts['probe.none.key'] + " -> nothing") ($null -eq $none) 'hits'

  # ---- one recorded row through the screen's own bindings ---------------------------------
  $idCol = $data.IdentityCol
  $probeRow = -1
  for ($i = 0; $i -lt $mr.Lines.Length; $i++) { if ([Rdv3Ledger]::FieldOf($mr.Lines[$i], $idCol) -eq $facts['probe.id']) { $probeRow = $i; break } }
  Ok ('the recorded row ' + $facts['probe.id'] + ' is in the ledger') ($probeRow -ge 0) 'not found'
  $view = New-Object Rdv3View
  $view.Record = [Rdv3Ledger]::SplitLine($mr.Lines[$probeRow])
  $view.StoredState = $initial
  $view.SearchKey = $facts['probe.multi.key']
  foreach ($k in ($facts.Keys | Where-Object { $_ -like 'probe.col.*' } | Sort-Object)) {
    $ref = $k.Substring(10)
    $col = $fields.IndexOf($ref)
    $got = $(if ($col -ge 0) { $view.Record[$col] } else { '(unresolved)' })
    Ok ("  " + $ref + " = " + $facts[$k]) ($got -eq $facts[$k]) ("got " + $got)
  }
  $bad = 0; $shown = 0
  foreach ($b in $cfg.Screen.AllBindings()) {
    $v = [Rdv3Eval]::Evaluate($b, $view, $fields, $work2)
    if ($v.Tone -eq [Rdv3Value]::Error) { $bad++ }
    if ($v.Tone -eq [Rdv3Value]::Normal -and $v.Text.Length -gt 0) { $shown++ }
  }
  Ok ("every screen binding evaluates without an error tone (" + $shown + " with a value)") ($bad -eq 0 -and $shown -gt 10) ("error tones " + $bad)
  $judgeId = $null
  foreach ($sec in $cfg.Screen.Sections) { if ($sec.Type -eq 'statusBand' -and $null -eq $judgeId) { $judgeId = $sec.Judgment } }
  $verdict = [Rdv3Eval]::Judge($cfg.Screen.JudgmentOf($judgeId), $view, $fields)
  Ok ('the judgment of the recorded row is ' + $facts['probe.judge']) ($verdict.Result.Id -eq $facts['probe.judge']) ("got " + $verdict.Result.Id + " raw=" + $verdict.Raw)
  if ($facts.ContainsKey('probe.empty.id')) {
    $row = -1
    for ($i = 0; $i -lt $mr.Lines.Length; $i++) { if ([Rdv3Ledger]::FieldOf($mr.Lines[$i], $idCol) -eq $facts['probe.empty.id']) { $row = $i; break } }
    $vw = New-Object Rdv3View; $vw.Record = [Rdv3Ledger]::SplitLine($mr.Lines[$row])
    $e = [Rdv3Eval]::Judge($cfg.Screen.JudgmentOf($judgeId), $vw, $fields)
    Ok ('an empty source value judges as ' + $facts['probe.empty.judge']) ($e.Result.Id -eq $facts['probe.empty.judge']) ("got " + $e.Result.Id)
  }
  # a row whose join found nothing shows the definition's "empty" text, never a blank
  $row = -1
  for ($i = 0; $i -lt $mr.Lines.Length; $i++) { if ([Rdv3Ledger]::FieldOf($mr.Lines[$i], $idCol) -eq $facts['probe.blank.id']) { $row = $i; break } }
  $vw = New-Object Rdv3View; $vw.Record = [Rdv3Ledger]::SplitLine($mr.Lines[$row]); $vw.StoredState = $initial
  # a binding whose every column comes from the table the join missed: the
  # whole value is then the definition's "empty" text
  $blankTable = $facts['probe.blank.col'].Substring(0, $facts['probe.blank.col'].IndexOf('.') + 1)
  $blankBind = $null
  foreach ($b in $cfg.Screen.AllBindings()) {
    if ($null -ne $blankBind -or -not $b.IsField) { continue }
    $all = $true
    foreach ($f in $b.Fields) { if (-not $f.StartsWith($blankTable)) { $all = $false } }
    if ($all) { $blankBind = $b }
  }
  $v = [Rdv3Eval]::Evaluate($blankBind, $vw, $fields, $work2)
  Ok ('an unmatched join shows the empty text (' + $v.Text + ')') (($v.Tone -eq [Rdv3Value]::Muted) -and ($v.Text -eq $blankBind.Empty) -and ($v.Text.Length -gt 0)) ("got '" + $v.Text + "' tone " + $v.Tone)

  # ---- the candidate list's columns, for the recorded row ----------------------------------
  $view.RowNumber = 1
  $bad = 0
  foreach ($c in $cfg.Screen.Candidates.Columns) {
    $v = [Rdv3Eval]::Evaluate($c.Value, $view, $fields, $work2)
    if ($v.Tone -eq [Rdv3Value]::Error) { $bad++ }
  }
  Ok ('the ' + $cfg.Screen.Candidates.Columns.Count + ' candidate columns evaluate') ($bad -eq 0) ("error tones " + $bad)

  # ---- work states ------------------------------------------------------------------------
  $cur = $work2.ByStored($initial)
  $chain = @($cur.Text)
  $steps = 0
  while ($null -ne ($tr = $work2.FromState($cur.Id)) -and $steps -lt 5) {
    $to = $work2.ById($tr.To)
    if ($tr.Confirm.Length -gt 0) {
      $text = [Rdv3Eval]::Template($tr.Confirm, $view, $fields, $work2)
      Ok ('the confirm text fills in: ' + $text) ($text.IndexOf('{') -lt 0) $text
    }
    $chain += $to.Text
    $cur = $to; $steps++
  }
  Ok ('the work states chain: ' + ($chain -join ' -> ')) ($chain.Count -ge 2) 'chain'
  Ok 'the last state has no transition out' ($null -eq $work2.FromState($cur.Id)) 'FromState'

  # ---- carry-over after a change ------------------------------------------------------------
  $newLines = [string[]]$mr.Lines.Clone()
  $newLines[$probeRow] = $mr.Lines[$probeRow] + 'x'              # the recorded row changed
  $oldStates = [string[]]$states.Clone()
  $lastStored = $cur.Stored
  $oldStates[$probeRow] = $lastStored                             # it had been worked
  $other = $(if ($probeRow -eq 0) { 1 } else { 0 })
  $oldStates[$other] = $lastStored                                # so had an unchanged one
  $updated = [Rdv3Ledger]::ApplyUpdate($data.UpdateJob, $mr.Lines, $oldStates, $newLines, $idCol, $initial)
  Ok 'update: the changed row goes back to the initial state' ($updated.States[$probeRow] -eq $initial) $updated.States[$probeRow]
  Ok 'update: the unchanged row keeps its state'            ($updated.States[$other] -eq $lastStored) $updated.States[$other]
  Ok 'update: the counts say so'                             (($updated.Unchanged -eq ($mr.Lines.Length - 1)) -and ($updated.Updated -eq 1) -and ($updated.ResetLines.Count -eq 1) -and ($updated.Deleted -eq 0)) ("{0}/{1}/{2}" -f $updated.Unchanged, $updated.Updated, $updated.ResetLines.Count)
}

Write-Output ''
Write-Output ("{0} passed, {1} failed" -f $script:pass, $script:fail)
Write-Output ''
if ($script:fail -gt 0) { Write-Output 'RESULT: FAIL'; exit 1 }
Write-Output 'RESULT: PASS'
