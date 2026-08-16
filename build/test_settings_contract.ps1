# ============================================================================
# test_settings_contract.ps1 -- what ReaderDataViewer.json is promised to MEAN.
#
#   powershell -File build\test_settings_contract.ps1
#
# The two geometry checks (test_ui_geometry.ps1, test_settings_geometry.ps1)
# prove the screen is the reference screen. Nothing proved what the settings
# file DOES, and that is where the promises are: docs\settings.md and the
# comments in the shipped src\app\config\ReaderDataViewer.json say a value out
# of range falls back on its own default, that the key length is the width of
# the key column, and that watch.targets is watched as written. Each case below
# is one of those sentences.
#
# The product is compiled from the shipping sources the same way the packer
# compiles it, and no window is ever created: this is the settings layer and the
# CSV reader, not the UI.
#
# ASCII only, so the file needs no BOM to survive Windows PowerShell 5.1.
# ============================================================================
[CmdletBinding()]
param([string] $Root = "")
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
if ([string]::IsNullOrEmpty($Root)) { $Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path) }

# ---- the product, compiled exactly as the packer compiles it ---------------
$sources = @('Rdv3Core.cs', 'Rdv3Index.cs', 'Rdv3Ledger.cs', 'Rdv3Xlsx.cs', 'Rdv3Jobs.cs',
             'Rdv3Watch.cs', 'Rdv3Text.cs', 'Rdv3Geom.cs', 'Rdv3SetGeom.cs', 'Rdv3Json.cs',
             'Rdv3Config.cs', 'Rdv3Uia.cs', 'Rdv3Ui.cs', 'Rdv3Settings.cs', 'Rdv3App.cs')
$usings = New-Object System.Collections.Specialized.OrderedDictionary
$bodies = New-Object System.Text.StringBuilder
foreach ($f in $sources) {
  $t = [IO.File]::ReadAllText((Join-Path $Root "src\app\csharp\$f"), [Text.Encoding]::UTF8)
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
Write-Output ''

$work = Join-Path $Root 'work\settings-contract'
if (Test-Path -LiteralPath $work) { Remove-Item -LiteralPath $work -Recurse -Force }
New-Item -ItemType Directory -Path $work | Out-Null

$script:pass = 0
$script:fail = 0
function Ok([string] $name, [bool] $cond, [string] $detail) {
  if ($cond) { $script:pass++; Write-Output ("  ok   {0}" -f $name) }
  else { $script:fail++; Write-Output ("  FAIL {0}  ({1})" -f $name, $detail) }
}
function Section([string] $t) { Write-Output ''; Write-Output $t }

function LoadJson([string] $name, [string] $body) {
  $p = Join-Path $work ($name + '.json')
  [IO.File]::WriteAllText($p, $body, (New-Object Text.UTF8Encoding($false)))
  return [Rdv3Config]::Load($p)
}
function NotesOf($c) { return ($c.Notes -join ' | ') }

# ---------------------------------------------------------------------------
Section 'a value out of range falls back on ITS OWN default, not on the limit'
# docs\settings.md: "a value that is out of range falls back on its own"
$c = LoadJson 'range' '{ "jobs": { "checkTimeoutMs": 0, "pumpMs": 999999 }, "key": { "length": 0 }, "watch": { "pollMs": 0 } }'
Ok 'jobs.checkTimeoutMs 0 -> 180000' ($c.CheckTimeoutMs -eq 180000) ("got " + $c.CheckTimeoutMs)
Ok 'jobs.pumpMs 999999 -> 1000'      ($c.PumpMs -eq 1000)           ("got " + $c.PumpMs)
Ok 'key.length 0 -> 8'               ($c.KeyLength -eq 8)           ("got " + $c.KeyLength)
Ok 'watch.pollMs 0 -> 40'            ($c.PollMs -eq 40)             ("got " + $c.PollMs)
Ok 'the reason is in the notes'      ((NotesOf $c) -match 'out of range') (NotesOf $c)

$c = LoadJson 'inrange' '{ "jobs": { "checkTimeoutMs": 1000 }, "key": { "length": 6 }, "watch": { "pollMs": 5 } }'
Ok 'a legal boundary is kept (1000)' ($c.CheckTimeoutMs -eq 1000) ("got " + $c.CheckTimeoutMs)
Ok 'a legal boundary is kept (5)'    ($c.PollMs -eq 5)            ("got " + $c.PollMs)
Ok 'key.length 6 is kept'            ($c.KeyLength -eq 6)         ("got " + $c.KeyLength)

# ---------------------------------------------------------------------------
Section 'watch.targets is watched AS WRITTEN'
$two = '{ "watch": { "targets": [' +
  '{ "enabled": true,  "name": "on",  "window": { "className": "Notepad" }, "field": { "controlTypes": ["Edit"] } },' +
  '{ "enabled": false, "name": "off", "window": { "className": "LobApp" },  "field": { "controlTypes": ["Edit"] } } ] } }'
$c = LoadJson 'mixed' $two
Ok 'a target turned off is still kept' ($c.Targets.Count -eq 2) ("count " + $c.Targets.Count)
if ($c.Targets.Count -eq 2) {
  Ok 'and keeps its enabled flag' (($c.Targets[0].Enabled -eq $true) -and ($c.Targets[1].Enabled -eq $false)) 'flags'
  # keeping a disabled target must not make it look watched: the status line and
  # the watcher both count IsWatchable, so 1-on-1-off is still ONE watched target
  # and the status line keeps naming it
  Ok 'the one that is off is not watchable' ($c.Targets[1].IsWatchable -eq $false) 'IsWatchable'
  Ok 'the one that is on is watchable'      ($c.Targets[0].IsWatchable -eq $true) 'IsWatchable'
  $on = @($c.Targets | Where-Object { $_.IsWatchable })
  Ok 'exactly one target is watched'        ($on.Count -eq 1) ("watched " + $on.Count)
  Ok 'and it is the one that was left on'   (($on.Count -eq 1) -and ($on[0].Name -eq 'on')) 'name'
  Ok 'the log line separates the two counts' (($c.Describe() -match 'targets=2') -and ($c.Describe() -match 'watched=1')) $c.Describe()
  Ok 'and marks the one that is off'         ($c.Describe() -match '\[off OFF:') $c.Describe()
}

$allOff = '{ "watch": { "targets": [' +
  '{ "enabled": false, "name": "a", "window": { "className": "Notepad" }, "field": { "controlTypes": ["Edit"] } } ] } }'
$c = LoadJson 'alloff' $allOff
Ok 'all off: the target is kept, not replaced' ($c.Targets.Count -eq 1) ("count " + $c.Targets.Count)
if ($c.Targets.Count -eq 1) {
  Ok 'all off: notepad is NOT revived' (($c.Targets[0].Enabled -eq $false) -and ($c.Targets[0].Window.ClassName -eq 'Notepad')) 'the one that was written'
}

$c = LoadJson 'empty' '{ "watch": { "targets": [] } }'
Ok 'an empty list means nothing is watched' ($c.Targets.Count -eq 0) ("count " + $c.Targets.Count)

$c = LoadJson 'nowatch' '{ "key": { "length": 8 } }'
Ok 'no watch member: the built-in target stands' ($c.Targets.Count -eq 1) ("count " + $c.Targets.Count)

# ---------------------------------------------------------------------------
Section 'a target that would match ANY window is not watched'
$blank = '{ "watch": { "targets": [ { "enabled": true, "name": "blank", "window": {}, "field": {} } ] } }'
$c = LoadJson 'blank' $blank
Ok 'window and field both blank -> dropped' ($c.Targets.Count -eq 0) ("count " + $c.Targets.Count)
Ok 'and the log says why' ((NotesOf $c) -match 'matches nothing in particular') (NotesOf $c)

$typo = '{ "watch": { "targets": [ { "enabled": true, "name": "typo",' +
  ' "window": { "className": "LobApp" }, "field": { "controlTypes": ["Edti"] } } ] } }'
$c = LoadJson 'typo' $typo
Ok 'a field named only by a misspelt controlType -> dropped' ($c.Targets.Count -eq 0) ("count " + $c.Targets.Count)
Ok 'and the log names the type it did not know' ((NotesOf $c) -match 'Edti') (NotesOf $c)

$good = '{ "watch": { "targets": [ { "enabled": true, "name": "good",' +
  ' "window": { "className": "LobApp" }, "field": { "controlTypes": ["Edit"] } } ] } }'
$c = LoadJson 'good' $good
Ok 'the same target spelt correctly is watched' ($c.Targets.Count -eq 1) ("count " + $c.Targets.Count)

$m = New-Object Rdv3Match
$m.ControlTypes = [string[]]@('Edti')
Ok 'IsEmpty: an unknown control type is not a constraint' ($m.IsEmpty -eq $true) 'IsEmpty'
$m.ControlTypes = [string[]]@('Edit')
Ok 'IsEmpty: a known control type is a constraint' ($m.IsEmpty -eq $false) 'IsEmpty'

# ---------------------------------------------------------------------------
Section 'what a running session may adopt, and what waits for the next start'
$a = [Rdv3Config]::Defaults()
$b = [Rdv3Config]::Defaults()
$b.KeyLength = 6; $b.KeyDigitsOnly = $false
$b.DataDir = 'D:\other'; $b.Ledger = 'D:\other\l.xlsx'; $b.Log = 'D:\other\r.log'
$b.PollMs = 77; $b.CandidateRowsShown = 3
$a.AdoptRuntimeFrom($b)
Ok 'runtime: pollMs is taken'            ($a.PollMs -eq 77) ("got " + $a.PollMs)
Ok 'runtime: candidateRowsShown taken'   ($a.CandidateRowsShown -eq 3) ("got " + $a.CandidateRowsShown)
# digitsOnly is only a rule about CHARACTERS. It never reaches the CSV reader or
# the index, so nothing built at start-up depends on it and it must NOT be made
# to wait for a restart the way key.length has to.
Ok 'runtime: key.digitsOnly IS taken'    ($a.KeyDigitsOnly -eq $false) ("got " + $a.KeyDigitsOnly)
Ok 'runtime: key.length is NOT taken'    ($a.KeyLength -eq 8) ("got " + $a.KeyLength)
Ok 'runtime: paths are NOT taken'        ($a.DataDir -eq 'data') ("got " + $a.DataDir)
$a.AdoptSavedFrom($b)
Ok 'saved: key.length is recorded'       ($a.KeyLength -eq 6) ("got " + $a.KeyLength)
Ok 'saved: dataDir is recorded'          ($a.DataDir -eq 'D:\other') ("got " + $a.DataDir)
Ok 'saved: ledger is recorded'           ($a.Ledger -eq 'D:\other\l.xlsx') ("got " + $a.Ledger)
Ok 'saved: log is recorded'              ($a.Log -eq 'D:\other\r.log') ("got " + $a.Log)

# a path edited, saved and edited again survives -- the second save must not
# write back the value the dialog was opened on
$p = Join-Path $work 'roundtrip.json'
$first = [Rdv3Config]::Defaults()
$first.DataDir = 'D:\one'
$null = $first.Save($p)
$live = [Rdv3Config]::Load($p)
$live.AdoptSavedFrom($first)
$second = $live.Clone()
$second.PollMs = 123
$null = $second.Save($p)
$after = [Rdv3Config]::Load($p)
Ok 'a saved path survives the next save' ($after.DataDir -eq 'D:\one') ("got " + $after.DataDir)
Ok 'and the next save carries its own change' ($after.PollMs -eq 123) ("got " + $after.PollMs)

# ---------------------------------------------------------------------------
Section 'the key length is the width of the key column'
function WriteCsv([string] $path, [int] $keyLen, [int] $rows) {
  $sb = New-Object System.Text.StringBuilder
  [void]$sb.Append('key1,key2')
  for ($f = 2; $f -lt 10; $f++) { [void]$sb.Append(',f' + $f) }
  [void]$sb.Append("`r`n")
  for ($i = 0; $i -lt $rows; $i++) {
    $k = ($i + 1).ToString().PadLeft($keyLen, '0')
    [void]$sb.Append($k).Append(',').Append($k)
    for ($f = 2; $f -lt 10; $f++) { [void]$sb.Append(',v' + $f) }
    [void]$sb.Append("`r`n")
  }
  [IO.File]::WriteAllText($path, $sb.ToString(), (New-Object Text.UTF8Encoding($false)))
}
$six = Join-Path $work 'six.csv'
$eight = Join-Path $work 'eight.csv'
WriteCsv $six 6 4
WriteCsv $eight 8 4

[Rdv3Spec]::KeyLength = 6
$t = [Rdv3Table]::Read($six, 'six')
Ok 'key.length 6 reads a 6 wide key column' ($t.Rows -eq 4) ("rows " + $t.Rows)
$ix = New-Object Rdv3Index $t
$rows = $null
$n = $ix.FindBytes($t.Buf, $t.KeyAt[2], [ref]$rows)
Ok 'and the index answers on the 6 wide key' ($n -eq 1) ("hits " + $n)
Ok 'with the row it was built from' (($n -eq 1) -and ($rows[0] -eq 2)) 'row'

$threw = $false
try { $null = [Rdv3Table]::Read($eight, 'eight') } catch { $threw = $true; $why = $_.Exception.Message }
Ok 'key.length 6 REFUSES an 8 wide key column' $threw 'it was accepted'
Ok 'and says what it wanted' ($threw -and ($why -match 'must be 6 bytes')) ("msg " + $why)

[Rdv3Spec]::KeyLength = 8
$t = [Rdv3Table]::Read($eight, 'eight')
Ok 'the default 8 still reads the shipped shape' ($t.Rows -eq 4) ("rows " + $t.Rows)
[Rdv3Spec]::KeyLength = [Rdv3Spec]::KeyLen

# ---------------------------------------------------------------------------
Write-Output ''
Write-Output ("{0} passed, {1} failed" -f $script:pass, $script:fail)
Write-Output ''
if ($script:fail -gt 0) { Write-Output 'RESULT: FAIL'; exit 1 }
Write-Output 'RESULT: PASS'
