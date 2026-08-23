# ============================================================================
# test_settings_contract.ps1 -- what settings.json is promised to MEAN.
#
#   powershell -File build\test_settings_contract.ps1
#
# The geometry checks prove the screen is the reference screen. This one
# proves what the file DOES: the shipped file loads and is the only definition
# there is; a file that is not right -- a typo, a value out of range, a name
# that refers to nothing -- does not load at all (no default, no half-applied
# file); the dialog writes only its three members and leaves every other byte;
# a value names a ledger column; a judgment that matches nothing is undefined
# and never OK; a stored state the definition does not know is refused; the
# CSV reader refuses by file and row what it cannot read as it is; the merge
# follows the definition; and the ledger file round-trips with its state column.
#
# Compiled from the shipping sources the way the packer compiles them; no
# window is ever created.
# ============================================================================
[CmdletBinding()]
param([string] $Root = "")
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
function WriteText([string] $name, [string] $body) {
  $p = Join-Path $work $name
  [IO.File]::WriteAllText($p, $body, (New-Object Text.UTF8Encoding($false)))
  return $p
}
$shippedPath = Join-Path $Root 'src\config\settings.json'
$shipped = [IO.File]::ReadAllText($shippedPath, [Text.Encoding]::UTF8)
# the shipped file with one piece of text replaced: the way a person breaks it
function Variant([string] $name, [string] $from, [string] $to) {
  if (-not $shipped.Contains($from)) { throw ("variant " + $name + ": the shipped text lacks: " + $from) }
  return WriteText ($name + '.json') ($shipped.Replace($from, $to))
}
function ErrorOf([scriptblock] $sb) {
  try { & $sb | Out-Null; return $null }
  catch { $e = $_.Exception; if ($e.InnerException) { $e = $e.InnerException }; return $e.Message }
}
# the file does not load, and the reason names what is wrong
function Refused([string] $name, [string] $path, [string] $pattern) {
  $m = ErrorOf { [Rdv3Config]::Load($path) }
  if ($null -eq $m) { Ok $name $false 'it loaded' } else { Ok $name ($m -match $pattern) $m }
}

# ===========================================================================
Section 'the shipped settings.json loads, and it is the only definition there is'
$cfg = [Rdv3Config]::Load($shippedPath)
Ok 'paths / search / watch / jobs are read' (($cfg.DataDir -eq 'data') -and ($cfg.KeyPattern -eq '^[0-9]{8}$') -and ($cfg.PollMs -eq 40) -and ($cfg.PumpMs -eq 1000) -and ($cfg.CandidateRowsShown -eq 100)) $cfg.Describe()
Ok 'one target, watched'                 (($cfg.Targets.Count -eq 1) -and $cfg.Targets[0].IsWatchable) $cfg.Describe()
Ok 'data: 3 tables, spine B, 2 joins'    (($cfg.Data.Tables.Count -eq 3) -and ($cfg.Data.Spine -eq 'B') -and ($cfg.Data.Joins.Count -eq 2)) $cfg.Data.Describe()
Ok 'data: 28 ledger columns'             ($cfg.Data.Columns.Count -eq 28) ("columns " + $cfg.Data.Columns.Count)
Ok 'identity is B.key2 (column 1), search is B.key1 (column 0)' (($cfg.Data.IdentityCol -eq 1) -and ($cfg.Data.SearchCol -eq 0)) $cfg.Data.Describe()
Ok 'the xlsx header is the CSV names'    (($cfg.Data.Head[0] -eq 'key1') -and ($cfg.Data.Head[2] -eq 'a_code') -and ($cfg.Data.Head[27] -eq 'c_remark')) ($cfg.Data.Head -join ',')
$types = @($cfg.Screen.Sections | ForEach-Object { $_.Type }) -join ','
Ok 'screen: the sections in reference order' ($types -eq 'keyPanel,columns,textBox,textBox,statusBand,statusBar') $types
Ok 'screen: judgment, 2 states, 10 candidate columns' (($cfg.Screen.Judgments.Count -eq 1) -and ($cfg.Screen.Work.States.Count -eq 2) -and ($cfg.Screen.Candidates.Columns.Count -eq 10)) $cfg.Screen.Describe()
Ok 'the key panel carries its three buttons' (($cfg.Screen.Sections[0].Buttons.Count -eq 3) -and ($cfg.Screen.Sections[0].Buttons[2].Action -eq 'workState')) 'buttons'
Ok 'the status bar carries the two actions'  (($cfg.Screen.Sections[5].Buttons.Count -eq 2) -and ($cfg.Screen.Sections[5].Buttons[0].Action -eq 'refreshLedger')) 'buttons'
Ok 'there is no built-in copy to fall back on' ($null -eq ([Rdv3Screen].GetMethod('Defaults')) -and $null -eq ([Rdv3Config].GetMethod('Defaults'))) 'Defaults() exists'

# ===========================================================================
Section 'a file that is not right does not load: no default, no half-applied file'
Refused 'a missing file'                         (Join-Path $work 'nowhere.json') 'does not exist'
Refused 'a syntax error, with the line'          (WriteText 'broken.json' ($shipped.Substring(0, $shipped.IndexOf('"jobs"')))) 'line [0-9]+'
Refused 'a typo in a member name (polMs)'        (Variant 'typo' '"pollMs": 40' '"polMs": 40') 'polMs.*not a member'
Refused 'a value out of range (pollMs 0)'        (Variant 'range' '"pollMs": 40' '"pollMs": 0') 'pollMs.*out of range'
Refused 'a value of the wrong kind'              (Variant 'kind' '"pollMs": 40' '"pollMs": "40"') 'pollMs.*must be a number'
Refused 'a whole number written as a fraction'   (Variant 'frac' '"pumpMs": 1000' '"pumpMs": 1000.5') 'pumpMs.*whole number'
Refused 'an unusable search pattern'             (Variant 'regex' '"pattern": "^[0-9]{8}$"' '"pattern": "([0-9]"') 'pattern.*regular expression'
Refused 'a member written twice'                 (Variant 'twice' '"pumpMs": 1000' '"pumpMs": 1000, "pumpMs": 2000') 'written twice'
Refused 'the old schema'                         (Variant 'schema' '"schema": 2' '"schema": 1') 'schema'
Refused 'a missing top-level part'               (Variant 'nojobs' '"jobs": {' '"jobz": {') 'jobz|jobs'
Refused 'a target that matches nothing'          (Variant 'blankwin' '"window": { "className": "Notepad", "scope": "children" }' '"window": { }') 'matches nothing'
Refused 'a control type UI Automation lacks'     (Variant 'ctype' '["Document", "Edit"]' '["Document", "Edti"]') 'Edti'
Refused 'data: a spine that is not a table'      (Variant 'spine' '"spine": "B"' '"spine": "X"') 'spine.*not one of the tables'
Refused 'data: a join to an unknown table'       (Variant 'jointbl' '{ "table": "A", "on": "key1" }' '{ "table": "Z", "on": "key1" }') 'not one of the tables'
Refused 'data: a table neither spine nor joined' (Variant 'unused' '{ "table": "C", "on": "key2" }' '{ "table": "A", "on": "key2" }') 'joined twice|neither the spine'
Refused 'data: a column written without its table' (Variant 'noref' '"A.a_code",' '"a_code",') '<table>.<column>'
Refused 'data: a ledger column listed twice'     (Variant 'dupcol' '"A.a_code",' '"A.a_code", "A.a_code",') 'listed twice'
Refused 'data: the identity missing from the ledger' (Variant 'noid' '"B.key1", "B.key2",' '"B.key1",') 'B.key2'
Refused 'data: a search column outside the ledger' (Variant 'search' '"search": "B.key1"' '"search": "B.b_unit2"') 'B.b_unit2'
Refused 'data: an encoding the machine lacks'    (Variant 'enc' '"encoding": "utf-8"' '"encoding": "klingon"') 'encoding'
Refused 'screen: a value naming a column the ledger lacks' (Variant 'col' '{ "field": "A.a_name" }' '{ "field": "A.a_phone" }') 'A.a_phone.*data.ledger.columns'
Refused 'screen: a section type nobody knows'    (Variant 'wizard' '"type": "textBox", "title": "メモ' '"type": "wizard", "title": "摘要') 'type.*must be one of'
Refused 'screen: a band naming a missing judgment' (Variant 'judg' '"judgment": "status1"' '"judgment": "status9"') 'status9.*not defined'
Refused 'screen: a transition to a state that is not there' (Variant 'trans' '"from": "todo", "to": "done",' '"from": "todo", "to": "dne",') 'dne.*not a state'
Refused 'screen: an initial state that is not there' (Variant 'init' '"initial": "todo"' '"initial": "nope"') 'nope.*not a state'
Refused 'screen: a rule naming a result with no entry' (Variant 'res' '"result": "ng" }' '"result": "bad" }') 'bad.*no entry'
Refused 'screen: a rule claiming the built-in result' (Variant 'builtin' '"result": "ng" }' '"result": "undefined" }') 'built-in result'
Refused 'screen: a confirm text naming a column the ledger lacks' (Variant 'tmpl' '{B.key2}' '{B.nope}') 'B.nope'
Refused 'screen: an app value the program does not provide' (Variant 'state' '{ "state": "searchKey" }' '{ "state": "moonPhase" }') 'moonPhase'
Refused 'screen: a look that is not a look'      (Variant 'look' '"look": "accent", "stored": "TRUE"' '"look": "shiny", "stored": "TRUE"') 'look.*shiny'

# ===========================================================================
Section 'search.pattern decides what a number is'
$c = [Rdv3Config]::Load((Variant 'six' '"pattern": "^[0-9]{8}$"' '"pattern": "^[0-9]{6}$"'))
Ok 'the pattern is taken as written' ($c.KeyPattern -eq '^[0-9]{6}$') $c.KeyPattern
Ok '6 digits is a key'              ($c.IsKey('123456')) 'IsKey'
Ok '8 digits is not'                (-not $c.IsKey('12345678')) 'IsKey'
Ok 'letters are not'                (-not $c.IsKey('12345A')) 'IsKey'
Ok 'the shipped pattern: 8 digits'  ($cfg.IsKey('00021001') -and -not $cfg.IsKey('0002100')) 'IsKey'

# ===========================================================================
Section 'watch.targets is watched AS WRITTEN'
$two = '{ "enabled": true,  "name": "on",  "window": { "className": "Notepad" }, "path": [], "field": { "controlTypes": ["Edit"] }, "read": "value" },' +
       '{ "enabled": false, "name": "off", "window": { "className": "LobApp" },  "path": [], "field": { "controlTypes": ["Edit"] }, "read": "value" }'
$oneTarget = $shipped.Substring($shipped.IndexOf('"targets": [') + 12, $shipped.IndexOf('"jobs"') - $shipped.IndexOf('"targets": [') - 12)
$oneTarget = $oneTarget.Substring(0, $oneTarget.LastIndexOf(']'))
$c = [Rdv3Config]::Load((Variant 'mixed' $oneTarget ("`r`n      " + $two + "`r`n    ")))
Ok 'a target turned off is still kept' ($c.Targets.Count -eq 2) ("count " + $c.Targets.Count)
if ($c.Targets.Count -eq 2) {
  Ok 'and keeps its enabled flag' (($c.Targets[0].Enabled -eq $true) -and ($c.Targets[1].Enabled -eq $false)) 'flags'
  $on = @($c.Targets | Where-Object { $_.IsWatchable })
  Ok 'exactly one target is watched'        ($on.Count -eq 1) ("watched " + $on.Count)
  Ok 'the log line separates the two counts' (($c.Describe() -match 'targets=2') -and ($c.Describe() -match 'watched=1')) $c.Describe()
}
$c = [Rdv3Config]::Load((Variant 'empty' $oneTarget ' '))
Ok 'an empty list means nothing is watched' ($c.Targets.Count -eq 0) ("count " + $c.Targets.Count)

# ===========================================================================
Section 'the dialog writes only paths / search / watch; every other byte stays'
$p = Join-Path $work 'site.json'
[IO.File]::WriteAllText($p, $shipped, (New-Object Text.UTF8Encoding($false)))
$live = [Rdv3Config]::Load($p)
$edited = $live.Clone()
$edited.PollMs = 77; $edited.KeyPattern = '^[A-Z0-9]{10}$'; $edited.DataDir = 'D:\one'
$edited.Targets[0].Enabled = $false
$err = $edited.Save($p)
Ok 'the save succeeds'                    ($null -eq $err) $err
$after = [IO.File]::ReadAllText($p, [Text.Encoding]::UTF8)
$reload = [Rdv3Config]::Load($p)
Ok 'the new values come back'             (($reload.PollMs -eq 77) -and ($reload.KeyPattern -eq '^[A-Z0-9]{10}$') -and ($reload.DataDir -eq 'D:\one') -and (-not $reload.Targets[0].Enabled)) $reload.Describe()
function Part([string] $text, [string] $from, [string] $to) { $a = $text.IndexOf($from); $b = $text.IndexOf($to, $a); return $text.Substring($a, $b - $a) }
Ok 'the header comment is untouched'      ($after.StartsWith($shipped.Substring(0, $shipped.IndexOf('{')))) 'header'
Ok 'the comment above "watch" is untouched' ($after.Contains('// WHAT TO WATCH.') -and $after.Contains('// point at -- far quicker')) 'comment'
Ok 'the "jobs" text is byte for byte'     ((Part $after '"jobs"' '// THE DATA') -eq (Part $shipped '"jobs"' '// THE DATA')) 'jobs'
Ok 'the "data" text is byte for byte'     ((Part $after '"data": {' '// THE SCREEN') -eq (Part $shipped '"data": {' '// THE SCREEN')) 'data'
Ok 'the "screen" text is byte for byte'   ($after.Substring($after.IndexOf('"screen": {')) -eq $shipped.Substring($shipped.IndexOf('"screen": {'))) 'screen'
Ok 'the file still ends the way it began' ($after.EndsWith($shipped.Substring($shipped.Length - 4))) 'tail'
$edited2 = $reload.Clone(); $edited2.StableMs = 99
$null = $edited2.Save($p)
$again = [Rdv3Config]::Load($p)
Ok 'a second save keeps the first one''s change' (($again.PollMs -eq 77) -and ($again.StableMs -eq 99) -and ($again.DataDir -eq 'D:\one')) $again.Describe()
# a file broken by hand while the app runs is not written over
$before = [IO.File]::ReadAllText($p, [Text.Encoding]::UTF8)
[IO.File]::WriteAllText($p, $before.Replace('"pumpMs": 1000', '"pumpMs": "soon"'), (New-Object Text.UTF8Encoding($false)))
$broken = [IO.File]::ReadAllText($p, [Text.Encoding]::UTF8)
$err = $again.Save($p)
Ok 'a save onto a file that no longer loads is refused' (($null -ne $err) -and ($err -match 'pumpMs')) $err
Ok 'and the file is left as it was'       ([IO.File]::ReadAllText($p, [Text.Encoding]::UTF8) -eq $broken) 'file changed'

# ===========================================================================
Section 'what a running session may adopt, and what waits for the next start'
$a = New-Object Rdv3Config
$b = New-Object Rdv3Config
$b.KeyPattern = '^[0-9]{6}$'
$b.DataDir = 'D:\other'; $b.Ledger = 'D:\other\l.xlsx'; $b.Log = 'D:\other\r.log'
$b.PollMs = 77; $b.CandidateRowsShown = 3
$a.AdoptRuntimeFrom($b)
Ok 'runtime: pollMs is taken'          ($a.PollMs -eq 77) ("got " + $a.PollMs)
Ok 'runtime: candidateRowsShown taken' ($a.CandidateRowsShown -eq 3) ("got " + $a.CandidateRowsShown)
Ok 'runtime: search.pattern IS taken'  ($a.KeyPattern -eq '^[0-9]{6}$') ("got " + $a.KeyPattern)
Ok 'runtime: paths are NOT taken'      ($a.DataDir -eq 'data') ("got " + $a.DataDir)
$a.AdoptSavedFrom($b)
Ok 'saved: the paths are recorded'     ($a.DataDir -eq 'D:\other' -and $a.Ledger -eq 'D:\other\l.xlsx' -and $a.Log -eq 'D:\other\r.log') 'paths'

# ===========================================================================
Section 'a value names a ledger column; the ledger answers, or the screen says it cannot'
$fields = New-Object Rdv3Fields (,$cfg.Data.ColumnRefs)
Ok '"A.a_name" is ledger column 3'   ($fields.IndexOf('A.a_name') -eq 3) ("got " + $fields.IndexOf('A.a_name'))
Ok '"B.key2" is ledger column 1'     ($fields.IndexOf('B.key2') -eq 1) ("got " + $fields.IndexOf('B.key2'))
Ok '"C.c_remark" is ledger column 27' ($fields.IndexOf('C.c_remark') -eq 27) ("got " + $fields.IndexOf('C.c_remark'))
Ok 'a bare name resolves too'        ($fields.IndexOf('b_status') -eq 16) ("got " + $fields.IndexOf('b_status'))
Ok 'a name the ledger lacks does not' ($fields.IndexOf('A.a_phone') -eq -1) ("got " + $fields.IndexOf('A.a_phone'))

$line = @('00021001','00089897','A81490','CUSTOMER-0021001','B2','20230311','7435696','0.5924','N','D737','NOTE-1',
          'SL00021001','20230311','597','44398','26505606','HOLD','001','',
          'IT347201','MAKER-2758','MECH','591852','43525','L60084','LOT84635','20250808','RMK-449812') -join "`t"
$view = New-Object Rdv3View
$work2 = $cfg.Screen.Work
function Bind([string] $json) { return [Rdv3Bind]::Read([Rdv3Json]::Parse($json)) }
$v = [Rdv3Eval]::Evaluate((Bind '{ "field": "A.a_name" }'), $view, $fields, $work2)
Ok 'no record: blank'                 (($v.Text -eq '') -and ($v.Tone -eq [Rdv3Value]::Muted)) $v.Text
$view.Record = [Rdv3Ledger]::SplitLine($line)
$view.StoredState = 'FALSE'
$v = [Rdv3Eval]::Evaluate((Bind '{ "field": "A.a_name" }'), $view, $fields, $work2)
Ok 'a column: its raw value'          (($v.Text -eq 'CUSTOMER-0021001') -and ($v.Tone -eq [Rdv3Value]::Normal)) $v.Text
$v = [Rdv3Eval]::Evaluate((Bind '{ "field": "B.b_memo" }'), $view, $fields, $work2)
Ok 'an empty column: N/A, muted'      (($v.Text -eq 'N/A') -and ($v.Tone -eq [Rdv3Value]::Muted)) $v.Text
$v = [Rdv3Eval]::Evaluate((Bind '{ "field": "B.b_memo", "empty": "(なし)" }'), $view, $fields, $work2)
Ok 'the empty text is the definition''s' ($v.Text -eq '(なし)') $v.Text
$v = [Rdv3Eval]::Evaluate((Bind '{ "field": "A.a_phone" }'), $view, $fields, $work2)
Ok 'a column the ledger lacks: shown as unresolved, error tone' (($v.Text -eq [Rdv3Text]::FieldUnresolved) -and ($v.Tone -eq [Rdv3Value]::Error)) $v.Text
$v = [Rdv3Eval]::Evaluate((Bind '{ "fields": ["A.a_rate", "A.a_flag"], "joiner": " / " }'), $view, $fields, $work2)
Ok 'several columns, joined'          ($v.Text -eq '0.5924 / N') $v.Text
$v = [Rdv3Eval]::Evaluate((Bind '{ "field": "A.a_amount", "format": { "kind": "number", "group": true } }'), $view, $fields, $work2)
Ok 'format number: thousands'         ($v.Text -eq '7,435,696') $v.Text
$v = [Rdv3Eval]::Evaluate((Bind '{ "field": "A.a_date", "format": { "kind": "date", "from": "yyyyMMdd", "to": "yyyy-MM-dd" } }'), $view, $fields, $work2)
Ok 'format date: reshaped'            ($v.Text -eq '2023-03-11') $v.Text
$v = [Rdv3Eval]::Evaluate((Bind '{ "field": "B.b_status", "format": { "kind": "date" } }'), $view, $fields, $work2)
Ok 'a value that is not a date stays as it is' ($v.Text -eq 'HOLD') $v.Text
$view.SearchKey = '00021001'
$v = [Rdv3Eval]::Evaluate((Bind '{ "state": "searchKey" }'), $view, $fields, $work2)
Ok 'an app value by name'             ($v.Text -eq '00021001') $v.Text
$m = ErrorOf { Bind '{ "field": "A.a_name", "state": "searchKey" }' }
Ok 'a value naming both a field and a state is refused' ($m -match 'exactly one') $m
$m = ErrorOf { Bind '{ "field": "A.a_name", "format": { "kind": "money" } }' }
Ok 'a format kind nobody knows is refused' ($m -match 'kind') $m

# ===========================================================================
Section 'a judgment is decided from the raw column, and nothing is OK by default'
$j = $cfg.Screen.JudgmentOf('status1')
function JudgeOf([string] $status) {
  $vw = New-Object Rdv3View
  $l = $line -replace "`tHOLD`t", ("`t" + $status + "`t")
  $vw.Record = [Rdv3Ledger]::SplitLine($l)
  return [Rdv3Eval]::Judge($j, $vw, $fields)
}
Ok 'DONE -> ok'                       ((JudgeOf 'DONE').Result.Id -eq 'ok') (JudgeOf 'DONE').Result.Id
Ok 'OPEN -> ok'                       ((JudgeOf 'OPEN').Result.Id -eq 'ok') (JudgeOf 'OPEN').Result.Id
Ok 'HOLD -> ng'                       ((JudgeOf 'HOLD').Result.Id -eq 'ng') (JudgeOf 'HOLD').Result.Id
Ok 'VOID -> ng'                       ((JudgeOf 'VOID').Result.Id -eq 'ng') (JudgeOf 'VOID').Result.Id
Ok 'an unknown value -> undefined'    ((JudgeOf 'WEIRD').Result.Id -eq 'undefined') (JudgeOf 'WEIRD').Result.Id
Ok 'an empty value -> undefined'      ((JudgeOf '').Result.Id -eq 'undefined') (JudgeOf '').Result.Id
Ok 'undefined wears its own look'     ((JudgeOf 'WEIRD').Result.Look -eq 'undefined') (JudgeOf 'WEIRD').Result.Look
$noRec = [Rdv3Eval]::Judge($j, (New-Object Rdv3View), $fields)
Ok 'no record -> nothing to judge'    ($null -eq $noRec.Result) 'Result'
function JudgmentOf([string] $json) { return [Rdv3Judgment]::Read('x', [Rdv3Json]::Parse($json)) }
$jErr = JudgmentOf '{ "source": { "field": "B.b_phase" }, "rules": [ { "equals": ["A"], "result": "ok" } ], "results": { "ok": { "text": "OK" } } }'
$vw = New-Object Rdv3View; $vw.Record = [Rdv3Ledger]::SplitLine($line)
$e = [Rdv3Eval]::Judge($jErr, $vw, $fields)
Ok 'a source column the ledger lacks -> error, not OK' (($e.Result.Id -eq 'error') -and ($e.Result.Look -eq 'error')) $e.Result.Id
$jEmpty = JudgmentOf '{ "source": { "field": "B.b_memo" }, "rules": [ { "empty": true, "result": "ng" } ], "results": { "ng": { "text": "NG", "look": "ng" } } }'
$e = [Rdv3Eval]::Judge($jEmpty, $vw, $fields)
Ok 'an "empty" rule catches a blank column' ($e.Result.Id -eq 'ng') $e.Result.Id
$jPat = JudgmentOf '{ "source": { "field": "A.a_grade" }, "rules": [ { "pattern": "^B", "result": "ok" } ], "results": { "ok": { "text": "OK" } } }'
$e = [Rdv3Eval]::Judge($jPat, $vw, $fields)
Ok 'a "pattern" rule matches a regular expression' ($e.Result.Id -eq 'ok') $e.Result.Id
$m = ErrorOf { JudgmentOf '{ "source": { "field": "B.b_status" }, "rules": [ { "result": "ok" } ], "results": { "ok": { "text": "OK" } } }' }
Ok 'a rule without a condition is refused' ($m -match 'no condition') $m
$m = ErrorOf { JudgmentOf '{ "source": { "field": "B.b_status" }, "rules": [ { "pattern": "([", "result": "ok" } ], "results": { "ok": { "text": "OK" } } }' }
Ok 'an unusable pattern is refused'        ($m -match 'regular expression') $m

# ===========================================================================
Section 'the work state is a stored string the definition names; an unknown one is refused'
$w = $cfg.Screen.Work
Ok 'FALSE is todo, TRUE is done'      (($w.ByStored('FALSE').Id -eq 'todo') -and ($w.ByStored('TRUE').Id -eq 'done')) 'ByStored'
Ok 'the initial state is todo'        ($w.InitialStored -eq 'FALSE') $w.InitialStored
Ok 'todo moves to done'               ($w.FromState('todo').To -eq 'done') 'FromState'
Ok 'done moves nowhere'               ($null -eq $w.FromState('done')) 'FromState'
Ok 'MAYBE is no state'                ($null -eq $w.ByStored('MAYBE')) 'ByStored'
$vw = New-Object Rdv3View; $vw.Record = [Rdv3Ledger]::SplitLine($line); $vw.StoredState = 'MAYBE'
$v = [Rdv3Eval]::WorkStateValue($vw, $w, $false)
Ok 'and is shown as itself, in the error tone' (($v.Text -eq 'MAYBE') -and ($v.Tone -eq [Rdv3Value]::Error)) $v.Text
$vw.StoredState = 'TRUE'; $vw.Saving = $true
$v = [Rdv3Eval]::WorkStateValue($vw, $w, $false)
Ok 'a save in flight says so'         ($v.Text -eq ('処理済' + [Rdv3Text]::SavingSuffix)) $v.Text
$t = [Rdv3Eval]::Template($w.Transitions[0].Confirm, $vw, $fields, $w)
Ok 'the confirm text is filled in'    ($t -eq '表示中のレコード (番号2 = 00089897) を処理済にします。よろしいですか?') $t
$t = [Rdv3Eval]::Template('{state} / {key} / {A.a_grade}', $vw, $fields, $w)
Ok '{state}, {key} and columns fill'  ($t -eq '処理済 (保存中...) /  / B2') $t
$three = [Rdv3WorkState]::Read([Rdv3Json]::Parse('{ "store": { "column": "s" }, "states": [ { "id": "a", "text": "A", "stored": "1" }, { "id": "b", "text": "B", "stored": "2" }, { "id": "c", "text": "C", "stored": "3" } ], "initial": "a", "transitions": [ { "from": "a", "to": "b" }, { "from": "b", "to": "c" }, { "from": "c", "to": "a" } ] }'))
Ok 'three states with a cycle load as written' (($three.States.Count -eq 3) -and ($three.FromState('c').To -eq 'a')) $three.States.Count
$m = ErrorOf { [Rdv3WorkState]::Read([Rdv3Json]::Parse('{ "store": { "column": "s" }, "states": [ { "id": "a", "stored": "1" }, { "id": "b", "stored": "1" } ], "initial": "a" }')) }
Ok 'two states with one stored value are refused' ($m -match 'used twice') $m

# ===========================================================================
Section 'carry-over across a rebuild keeps a stored state only for identical content'
$old = @("k1`tK2A`tx", "k1`tK2B`ty", "k1`tK2C`tz")
$oldStates = @('TRUE', 'TRUE', 'FALSE')
$new = @("k1`tK2A`tx", "k1`tK2B`tCHANGED", "k1`tK2D`tnew")
$stats = New-Object Rdv3Ledger+CarryStats
$ns = [Rdv3Ledger]::CarryStates($old, $oldStates, $new, 1, 'FALSE', $stats)
Ok 'identical content keeps TRUE'     ($ns[0] -eq 'TRUE') $ns[0]
Ok 'changed content resets'           ($ns[1] -eq 'FALSE') $ns[1]
Ok 'a new row starts at the initial'  ($ns[2] -eq 'FALSE') $ns[2]
Ok 'the counts say what happened'     (($stats.Carried -eq 1) -and ($stats.Reset -eq 1) -and ($stats.New -eq 1) -and ($stats.Dropped -eq 1)) ("{0}/{1}/{2}/{3}" -f $stats.Carried, $stats.Reset, $stats.New, $stats.Dropped)

# ===========================================================================
Section 'the CSV reader is strict: what it refuses, it refuses by file and row'
$utf8 = New-Object Text.UTF8Encoding($false)
function WriteCsv([string] $path, [int] $keyLen, [int] $rows, [int] $cols) {
  $sb = New-Object System.Text.StringBuilder
  [void]$sb.Append('key1,key2')
  for ($f = 2; $f -lt $cols; $f++) { [void]$sb.Append(',f' + $f) }
  [void]$sb.Append("`r`n")
  for ($i = 0; $i -lt $rows; $i++) {
    $k = ($i + 1).ToString().PadLeft($keyLen, '0')
    [void]$sb.Append($k).Append(',').Append($k)
    for ($f = 2; $f -lt $cols; $f++) { [void]$sb.Append(',v' + $f) }
    [void]$sb.Append("`r`n")
  }
  [IO.File]::WriteAllText($path, $sb.ToString(), $utf8)
}
function Csv([string] $name, [string] $body) { return WriteText ($name + '.csv') $body }
$six = Join-Path $work 'six.csv'; WriteCsv $six 6 4 10
$eight = Join-Path $work 'eight.csv'; WriteCsv $eight 8 4 10
$t6 = [Rdv3Table]::Read($six, 'six', $utf8, 'key1')
Ok 'a 6-wide key column reads as 6'  ($t6.KeyLen -eq 6 -and $t6.Rows -eq 4) ("keylen " + $t6.KeyLen)
$t8 = [Rdv3Table]::Read($eight, 'eight', $utf8, 'key2')
Ok 'an 8-wide key column reads as 8' ($t8.KeyLen -eq 8) ("keylen " + $t8.KeyLen)
Ok 'the key may be any column'       ($t8.KeyCol -eq 1) ("keycol " + $t8.KeyCol)
Ok 'the header names are kept'       ($t8.Head[1] -eq 'key2' -and $t8.Head.Length -eq 10) ($t8.Head -join ',')
$ix = New-Object Rdv3Index $t6
$rows = $null
$n = $ix.FindBytes($t6.Buf, $t6.KeyAt[2], 6, [ref]$rows)
Ok 'the index answers at the table''s width' (($n -eq 1) -and ($rows[0] -eq 2)) ("hits " + $n)
$n = $ix.FindBytes($t8.Buf, $t8.KeyAt[2], 8, [ref]$rows)
Ok 'and never at another width'      ($n -eq 0) ("hits " + $n)
$wide = Join-Path $work 'wide.csv'; WriteCsv $wide 8 3 12
$tw = [Rdv3Table]::Read($wide, 'wide', $utf8, 'key1')
Ok 'a table may have columns the ledger does not use' ($tw.Head.Length -eq 12 -and $tw.Rows -eq 3) ("cols " + $tw.Head.Length)
$m = ErrorOf { [Rdv3Table]::Read((Csv 'mixed' "key1,key2,f2`r`n000001,000001,a`r`n00000002,00000002,a`r`n"), 'mixed', $utf8, 'key1') }
Ok 'a row whose key differs in width is refused, by row' ($m -match 'mixed.csv.*3.*キー列 key1') $m
$m = ErrorOf { [Rdv3Table]::Read((Csv 'short' "key1,key2,f2`r`n000001,000001,a`r`n000002,000002`r`n"), 'short', $utf8, 'key1') }
Ok 'a row with fewer columns is refused, by row' ($m -match 'short.csv.*3.*列数') $m
$m = ErrorOf { [Rdv3Table]::Read((Csv 'long' "key1,key2,f2`r`n000001,000001,a,b`r`n"), 'long', $utf8, 'key1') }
Ok 'a row with more columns is refused'          ($m -match 'long.csv.*2.*列数') $m
$m = ErrorOf { [Rdv3Table]::Read((Csv 'quoted' "key1,key2,f2`r`n000001,000001,`"a, b`"`r`n"), 'quoted', $utf8, 'key1') }
Ok 'a quoted field is refused, not shifted'      ($m -match 'quoted.csv.*2.*引用符') $m
$m = ErrorOf { [Rdv3Table]::Read((Csv 'nokey' "key1,key2,f2`r`n000001,000001,a`r`n"), 'nokey', $utf8, 'key9') }
Ok 'a key column the header lacks is refused'    ($m -match 'nokey.csv.*key9') $m
$m = ErrorOf { [Rdv3Table]::Read((Csv 'duphead' "key1,key1,f2`r`n000001,000001,a`r`n"), 'duphead', $utf8, 'key1') }
Ok 'a header naming a column twice is refused'   ($m -match 'duphead.csv.*key1') $m
$m = ErrorOf { [Rdv3Table]::Read((Csv 'emptykey' "key1,key2,f2`r`n000001,000001,a`r`n,000002,a`r`n"), 'emptykey', $utf8, 'key1') }
Ok 'an empty key is refused'                     ($m -match 'emptykey.csv.*3') $m
$m = ErrorOf { [Rdv3Table]::Read((Csv 'headonly' "key1,key2,f2`r`n"), 'headonly', $utf8, 'key1') }
Ok 'a file with no data row is refused'          ($m -match 'headonly.csv') $m
$m = ErrorOf { New-Object Rdv3Index ([Rdv3Table]::Read((Csv 'dupkey' "key1,key2,f2`r`n000001,000001,a`r`n000002,000002,a`r`n000001,000003,a`r`n"), 'dupkey', $utf8, 'key1')) }
Ok 'a duplicate key is refused, naming both rows' ($m -match 'dupkey.csv.*000001.*2.*4') $m
$sjis = [Text.Encoding]::GetEncoding(932)
$sj = Join-Path $work 'sjis.csv'
[IO.File]::WriteAllText($sj, "key1,名前`r`n000001,漢字`r`n", $sjis)
$ts = [Rdv3Table]::Read($sj, 'sjis', $sjis, 'key1')
Ok 'a Shift_JIS table reads with its encoding'   (($ts.Head[1] -eq '名前') -and ($ts.Field(0, 1) -eq '漢字')) ($ts.Head -join ',')
$bom = Join-Path $work 'bom.csv'
[IO.File]::WriteAllText($bom, "key1,f1`r`n000001,a`r`n", (New-Object Text.UTF8Encoding($true)))
$tb = [Rdv3Table]::Read($bom, 'bom', $utf8, 'key1')
Ok 'a UTF-8 BOM is skipped'                      (($tb.Head[0] -eq 'key1') -and ([Rdv3Table]::ReadHead($bom, $utf8)[0] -eq 'key1')) ($tb.Head -join ',')

# ===========================================================================
Section 'the merge follows the definition: spine, joins, ledger columns'
$dd = Join-Path $work 'tiny'
New-Item -ItemType Directory -Path $dd | Out-Null
[IO.File]::WriteAllText((Join-Path $dd 'a.csv'), "key1,a_x`r`n0001,AX1`r`n0002,AX2`r`n", $utf8)
[IO.File]::WriteAllText((Join-Path $dd 'b.csv'), "key1,key2,b_x`r`n0001,K1,B1`r`n0001,K2,B2`r`n0009,K3,B3`r`n", $utf8)
[IO.File]::WriteAllText((Join-Path $dd 'c.csv'), "key2,c_x,c_y`r`nK1,C1,Y1`r`nK3,C3,Y3`r`n", $utf8)
$defJson = '{ "tables": { "A": { "file": "a.csv", "key": "key1" }, "B": { "file": "b.csv", "key": "key2" }, "C": { "file": "c.csv", "key": "key2" } },' +
           '  "spine": "B", "joins": [ { "table": "A", "on": "key1" }, { "table": "C", "on": "key2" } ],' +
           '  "ledger": { "search": "B.key1", "columns": [ "B.key2", "C.c_y", "A.a_x", "B.b_x", "B.key1" ] } }'
$def = [Rdv3Data]::Read([Rdv3Json]::Parse($defJson))
$mr = [Rdv3Ledger]::BuildFromCsv($def, $dd)
Ok 'one ledger row per spine row'        ($mr.Rows -eq 3) ("rows " + $mr.Rows)
Ok 'the columns come in ledger order'    ($mr.Lines[0] -eq "K1`tY1`tAX1`tB1`t0001") $mr.Lines[0]
Ok 'a row whose join finds nothing is blank there' ($mr.Lines[1] -eq "K2`t`tAX1`tB2`t0001") $mr.Lines[1]
Ok 'a spine key nobody has leaves A blank' ($mr.Lines[2] -eq "K3`tY3`t`tB3`t0009") $mr.Lines[2]
Ok 'the matched counts say so'           (($mr.Matched[0] -eq 2) -and ($mr.Matched[1] -eq 2)) ($mr.Matched -join ',')
Ok 'the xlsx header is the CSV names'    (($mr.Head -join ',') -eq 'key2,c_y,a_x,b_x,key1') ($mr.Head -join ',')
Ok 'identity and search columns follow'  (($def.IdentityCol -eq 0) -and ($def.SearchCol -eq 4)) ("{0}/{1}" -f $def.IdentityCol, $def.SearchCol)
$m = ErrorOf { [Rdv3Data]::Read([Rdv3Json]::Parse($defJson.Replace('"columns": [ "B.key2",', '"columns": ['))) }
Ok 'a ledger without the identity is refused' ($m -match 'B.key2') $m
$def2 = [Rdv3Data]::Read([Rdv3Json]::Parse($defJson.Replace('"C.c_y"', '"C.c_z"')))
$m = ErrorOf { [Rdv3Ledger]::BuildFromCsv($def2, $dd) }
Ok 'a column the CSV lacks stops the merge, naming the file' ($m -match 'c.csv.*c_z') $m
[IO.File]::WriteAllText((Join-Path $dd 'a.csv'), "key1,a_x`r`n0001,AX1`r`n0001,AX2`r`n", $utf8)
$m = ErrorOf { [Rdv3Ledger]::BuildFromCsv($def, $dd) }
Ok 'a duplicate key in a joined table stops the merge' ($m -match 'a.csv.*0001') $m

# ===========================================================================
Section 'the ledger file round-trips with its state column'
$head = $cfg.Data.Head
$xl = Join-Path $work 'round.xlsx'
$lines = @($line, ($line -replace '00089897', '00089898'))
[Rdv3Xlsx]::Write($xl, $head, '処理済み', $lines, @('TRUE', 'MAYBE'), 'test')
$outLines = $null; $outStates = $null
[Rdv3Xlsx]::Read($xl, $head, '処理済み', [ref]$outLines, [ref]$outStates)
Ok 'rows and content come back'      (($outLines.Length -eq 2) -and ($outLines[1] -eq $lines[1])) ("rows " + $outLines.Length)
Ok 'states come back verbatim'       (($outStates[0] -eq 'TRUE') -and ($outStates[1] -eq 'MAYBE')) ($outStates -join ',')
$m = ErrorOf { [Rdv3Xlsx]::Read($xl, $head, 'processed', [ref]$outLines, [ref]$outStates) }
Ok 'another state column heading is refused' ($m -match 'round.xlsx') $m
$m = ErrorOf { [Rdv3Xlsx]::Read($xl, $mr.Head, '処理済み', [ref]$outLines, [ref]$outStates) }
Ok 'another column layout is refused'        ($m -match 'round.xlsx') $m

# ---------------------------------------------------------------------------
Write-Output ''
Write-Output ("{0} passed, {1} failed" -f $script:pass, $script:fail)
Write-Output ''
if ($script:fail -gt 0) { Write-Output 'RESULT: FAIL'; exit 1 }
Write-Output 'RESULT: PASS'
