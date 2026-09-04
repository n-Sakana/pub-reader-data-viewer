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
Ok 'paths / search / watch / jobs are read' (($cfg.DataDir -eq 'data') -and ($cfg.KeyPattern -eq '^[0-9]{8}$') -and ($cfg.PollMs -eq 40) -and ($cfg.PumpMs -eq 1000) -and ($cfg.LockRetryMs -eq 250) -and ($cfg.MarkerPollMs -eq 3000) -and ($cfg.CandidateRowsShown -eq 100)) $cfg.Describe()
Ok 'one target, watched'                 (($cfg.Targets.Count -eq 1) -and $cfg.Targets[0].IsWatchable) $cfg.Describe()
Ok 'data: schema 3 has two named jobs'   (($cfg.Data.Jobs.Count -eq 2) -and ($cfg.Data.Jobs[0].Id -eq 'merge-ledger') -and ($cfg.Data.Jobs[1].Id -eq 'delete-listed-records')) $cfg.Data.Describe()
Ok 'data: update job compiles to spine B and 2 joins' (($cfg.Data.Tables.Count -eq 3) -and ($cfg.Data.Spine -eq 'B') -and ($cfg.Data.Joins.Count -eq 2)) $cfg.Data.Describe()
Ok 'data: 28 ledger columns'             ($cfg.Data.Columns.Count -eq 28) ("columns " + $cfg.Data.Columns.Count)
Ok 'the update is merge: add / update / keep' (($cfg.Data.UpdateJob.ApplyStep.Operation -eq 'merge') -and ($cfg.Data.UpdateJob.ApplyStep.SourceOnly -eq 'add') -and ($cfg.Data.UpdateJob.ApplyStep.Both -eq 'update') -and ($cfg.Data.UpdateJob.ApplyStep.TargetOnly -eq 'keep')) $cfg.Data.Describe()
Ok 'source and application columns are separated with an explicit change rule' (($cfg.Data.ApplicationColumns.Count -eq 1) -and $cfg.Data.WorkStateIsApplicationOwned -and ($cfg.Data.WorkStateOnSourceChange -eq 'reset')) (($cfg.Data.ApplicationColumns -join ',') + '/' + $cfg.Data.WorkStateOnSourceChange)
Ok 'identity is B.key2; exact search names B.key1' (($cfg.Data.IdentityCol -eq 1) -and ($cfg.Data.SearchCols.Count -eq 1) -and ($cfg.Data.SearchCols[0] -eq 0) -and ($cfg.Data.SearchMatch -eq 'exact')) $cfg.Data.Describe()
Ok 'the xlsx header is the CSV names'    (($cfg.Data.Head[0] -eq 'key1') -and ($cfg.Data.Head[2] -eq 'a_code') -and ($cfg.Data.Head[27] -eq 'c_remark')) ($cfg.Data.Head -join ',')
$types = @($cfg.Screen.Sections | ForEach-Object { $_.Type }) -join ','
Ok 'screen: the sections in reference order' ($types -eq 'keyPanel,columns,textBox,textBox,statusBand,sendBar,statusBar') $types
Ok 'screen: judgment, 2 states, 10 candidate columns' (($cfg.Screen.Judgments.Count -eq 1) -and ($cfg.Screen.Work.States.Count -eq 2) -and ($cfg.Screen.Candidates.Columns.Count -eq 10)) $cfg.Screen.Describe()
Ok 'screen: emphasis sizes come from the card definition' (($cfg.Screen.FontSize -eq 10) -and ($cfg.Screen.KeyValueFontSize -eq 15) -and ($cfg.Screen.JudgmentFontSize -eq 15) -and ($cfg.Screen.UnsearchedFontSize -eq 12)) $cfg.Screen.Describe()
Ok 'screen: automatic work-state change remains the default' ($cfg.Screen.Work.Trigger -eq 'automatic') $cfg.Screen.Work.Trigger
$typeSummary = @($cfg.Data.TypeOrder | ForEach-Object { $_.Ref + ':' + $_.Type + ':' + $_.Format }) -join ','
Ok 'data: the shipped definition declares 3 dates and 6 numbers' ($typeSummary -eq 'A.a_date:date:yyyyMMdd,B.b_date:date:yyyyMMdd,C.c_exp:date:yyyyMMdd,A.a_rate:number:,A.a_amount:number:,B.b_qty:number:,B.b_total:number:,C.c_price:number:,C.c_stock:number:') $typeSummary
$defaultKeyValidation = $cfg.Data.Tables[0].KeyValidation
$sampleBKeyValidation = $cfg.Data.Tables[1].KeyValidation
Ok 'data: omitted key validation stays strict; shipped B skips only blank identities' ($defaultKeyValidation.Ascii -and $defaultKeyValidation.FixedLength -and $defaultKeyValidation.Unique -and -not $defaultKeyValidation.SkipEmpty -and $sampleBKeyValidation.Ascii -and $sampleBKeyValidation.FixedLength -and $sampleBKeyValidation.Unique -and $sampleBKeyValidation.SkipEmpty) 'key rules changed'
Ok 'the key panel carries its three buttons' (($cfg.Screen.Sections[0].Buttons.Count -eq 3) -and ($cfg.Screen.Sections[0].Buttons[2].Action -eq 'workState')) 'buttons'
Ok 'the send band carries one direct action' (($cfg.Screen.Sections[5].Buttons.Count -eq 1) -and ($cfg.Screen.Sections[5].Buttons[0].Action -eq 'sendChanges')) 'send button'
Ok 'the status bar carries four actions'  (($cfg.Screen.Sections[6].Buttons.Count -eq 4) -and ($cfg.Screen.Sections[6].Buttons[0].Action -eq 'tableExport') -and ($cfg.Screen.Sections[6].Buttons[3].Action -eq 'settings')) 'buttons'
Ok 'update/delete buttons name their jobs directly' (($cfg.Screen.Sections[6].Buttons[1].Job -eq 'merge-ledger') -and ($cfg.Screen.Sections[6].Buttons[2].Job -eq 'delete-listed-records')) 'button jobs'
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
Refused 'the old schema'                         (Variant 'schema' '"schema": 3' '"schema": 2') 'schema'
Refused 'a missing top-level part'               (Variant 'nojobs' '"jobs": {' '"jobz": {') 'jobz|jobs'
Refused 'a target that matches nothing'          (Variant 'blankwin' '"window": { "className": "Notepad", "scope": "children" }' '"window": { }') 'matches nothing'
Refused 'a control type UI Automation lacks'     (Variant 'ctype' '["Document", "Edit"]' '["Document", "Edti"]') 'Edti'
Refused 'data: a job input names an unknown table' (Variant 'jobinput' '{ "table": "A" }' '{ "table": "Z" }') 'not one of the tables'
Refused 'data: a step consumes an unknown target' (Variant 'steptarget' '"target2": "A", "keys": ["B.key1", "A.key1"]' '"target2": "Z", "keys": ["B.key1", "A.key1"]') 'has not been defined'
Refused 'data: operations require matching value types' (Variant 'valuetype' '"operation": "extract", "target1": "target1", "target2": "target2"' '"operation": "append", "target1": "target1", "target2": "target2"') 'requires two tables'
Refused 'screen: a process button names no job'  (Variant 'buttonjob' '"job": "merge-ledger"' '"job": "missing-job"') 'missing-job.*data.jobs'
Refused 'data: a column written without its table' (Variant 'noref' '"A.a_code",' '"a_code",') '<table>.<column>'
Refused 'data: a ledger column listed twice'     (Variant 'dupcol' '"A.a_code",' '"A.a_code", "A.a_code",') 'listed twice'
Refused 'data: the identity missing from the ledger' (Variant 'noid' '"B.key1", "B.key2",' '"B.key1",') 'B.key2'
Refused 'data: a search column outside the ledger' (Variant 'search' '"columns": ["B.key1"], "match": "exact"' '"columns": ["B.b_unit2"], "match": "exact"') 'B.b_unit2'
Refused 'data: application ownership is explicit' (Variant 'owner' '{ "name": "workState", "onSourceChange": "reset" }' '') 'application.*no application-owned column'
Refused 'data: merge row destinations are strict' (Variant 'rowdest' '"targetOnly": "keep"' '"targetOnly": "drop"') 'targetOnly.*keep / delete'
Refused 'data: delete cannot target the whole ledger directly' (Variant 'unboundeddelete' '"operation": "delete", "target1": "ledger", "target2": "target"' '"operation": "delete", "target1": "ledger", "target2": "ledger"') 'requires a table and rows selected'
Refused 'data: an encoding the machine lacks'    (Variant 'enc' '"encoding": "utf-8"' '"encoding": "klingon"') 'encoding'
Refused 'data: a type must be date or number' (Variant 'badtype' '"A.a_date": { "type": "date", "format": "yyyyMMdd" }' '"A.a_date": { "type": "money", "format": "yyyyMMdd" }') 'type.*date / number'
Refused 'data: a date type requires its exact format' (Variant 'dateformat' '"B.b_date": { "type": "date", "format": "yyyyMMdd" }' '"B.b_date": { "type": "date" }') 'format.*required'
Refused 'data: a number type has no date format' (Variant 'numberformat' '"A.a_rate": { "type": "number" }' '"A.a_rate": { "type": "number", "format": "0" }') 'format.*only when type is date'
Refused 'data: a type entry keeps strict unknown-member checking' (Variant 'typemember' '"C.c_exp": { "type": "date", "format": "yyyyMMdd" }' '"C.c_exp": { "type": "date", "format": "yyyyMMdd", "guess": true }') 'guess.*not a member'
$tableA = '"A": { "label": "表A", "file": "tableA.csv", "key": "key1" },'
$relaxedTableA = '"A": { "label": "表A", "file": "tableA.csv", "key": "key1", "keyValidation": { "characters": "unicode", "length": "variable", "duplicates": "distinct", "empty": "skip" } },'
$relaxedKeyCfg = [Rdv3Config]::Load((Variant 'key-validation' $tableA $relaxedTableA))
$relaxedRules = $relaxedKeyCfg.Data.Tables[0].KeyValidation
Ok 'data: all four table key rules are read explicitly' ((-not $relaxedRules.Ascii) -and (-not $relaxedRules.FixedLength) -and (-not $relaxedRules.Unique) -and $relaxedRules.SkipEmpty) 'key rules were not read'
Refused 'data: key character validation is a closed choice' (Variant 'key-characters' $tableA ($relaxedTableA.Replace('"unicode"', '"binary"'))) 'characters.*ascii / unicode'
Refused 'data: key validation keeps strict unknown-member checking' (Variant 'key-member' $tableA ($relaxedTableA.Replace('"empty": "skip"', '"empty": "skip", "guess": true'))) 'guess.*not a member'
Refused 'screen: a work-state trigger is automatic or manual' (Variant 'trigger' '"store": { "column": "処理済み" }' '"trigger": "sometimes", "store": { "column": "処理済み" }') 'trigger.*automatic / manual'
Refused 'screen: an emphasis size stays in its declared range' (Variant 'emphasis' '"keyValueFontSize": 15' '"keyValueFontSize": 50') 'keyValueFontSize.*out of range'
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
$manualWork = [Rdv3Config]::Load((Variant 'manual-work' '"store": { "column": "処理済み" }' '"trigger": "manual", "store": { "column": "処理済み" }'))
Ok 'work-state automatic change can be disabled in JSON' ($manualWork.Screen.Work.Trigger -eq 'manual') $manualWork.Screen.Work.Trigger
$multiSearch = [Rdv3Config]::Load((Variant 'multisearch' '"columns": ["B.key1"], "match": "exact"' '"columns": ["B.key1", "B.key2"], "match": "contains"'))
Ok 'ledger search accepts several columns' (($multiSearch.Data.SearchCols.Count -eq 2) -and ($multiSearch.Data.SearchRefs[1] -eq 'B.key2')) ($multiSearch.Data.SearchRefs -join ',')
Ok 'ledger search accepts contains matching' ($multiSearch.Data.SearchMatch -eq 'contains') $multiSearch.Data.SearchMatch

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
Ok 'the paths / search / watch guide in the header is untouched' ($after.Contains('// paths の書き方:') -and $after.Contains('// watch の書き方:') -and $after.Contains('// 押すだけで、この定義を UI Automation の実物から作れます。')) 'comment'
Ok 'the "jobs" text is byte for byte'     ((Part $after '"jobs"' '// 入力表と統合台帳の定義。') -eq (Part $shipped '"jobs"' '// 入力表と統合台帳の定義。')) 'jobs'
Ok 'the "data" text is byte for byte'     ((Part $after '"data": {' '// 画面の定義。') -eq (Part $shipped '"data": {' '// 画面の定義。')) 'data'
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
Ok 'done moves back to todo'           ($w.FromState('done').To -eq 'todo') 'FromState'
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
$ns = [Rdv3Ledger]::CarryStates($old, $oldStates, $new, 1, 'FALSE', 'reset', $stats)
Ok 'identical content keeps TRUE'     ($ns[0] -eq 'TRUE') $ns[0]
Ok 'changed content resets'           ($ns[1] -eq 'FALSE') $ns[1]
Ok 'a new row starts at the initial'  ($ns[2] -eq 'FALSE') $ns[2]
Ok 'the counts say what happened'     (($stats.Carried -eq 1) -and ($stats.Reset -eq 1) -and ($stats.New -eq 1) -and ($stats.Dropped -eq 1)) ("{0}/{1}/{2}/{3}" -f $stats.Carried, $stats.Reset, $stats.New, $stats.Dropped)
$stats = New-Object Rdv3Ledger+CarryStats
$ns = [Rdv3Ledger]::CarryStates($old, $oldStates, $new, 1, 'FALSE', 'preserve', $stats)
Ok 'the configured preserve rule keeps state across changed source columns' ($ns[1] -eq 'TRUE') ($ns -join ',')

# ===========================================================================
Section 'merge grows one ledger, preserves application state, and resets changed rows'
$april = [string[]]@("APR`tID-APR-1`told-a", "APR`tID-APR-2`told-b")
$aprilStates = [string[]]@('TRUE', 'TRUE')
$may = [string[]]@("MAY`tID-APR-2`tchanged-b", "MAY`tID-MAY-1`tnew-c")
$grown = [Rdv3Ledger]::ApplyUpdate($cfg.Data.UpdateJob, $april, $aprilStates, $may, 1, 'FALSE')
Ok 'April plus May stays in one ledger' (($grown.Lines.Length -eq 3) -and ($grown.Lines[0] -eq $april[0]) -and ($grown.Lines[1] -eq $may[0]) -and ($grown.Lines[2] -eq $may[1])) ($grown.Lines -join '|')
Ok 'a destination-only April row and its mark remain' (($grown.States[0] -eq 'TRUE') -and ($grown.Kept -eq 1)) (($grown.States -join ',') + ' kept=' + $grown.Kept)
Ok 'changed source content returns a processed row to initial' (($grown.States[1] -eq 'FALSE') -and ($grown.Updated -eq 1) -and ($grown.ResetLines.Count -eq 1)) (($grown.States -join ',') + ' reset=' + $grown.ResetLines.Count)
Ok 'a May-only row is appended in the initial state' (($grown.States[2] -eq 'FALSE') -and ($grown.Added -eq 1)) (($grown.States -join ',') + ' added=' + $grown.Added)
$again = [Rdv3Ledger]::ApplyUpdate($cfg.Data.UpdateJob, $grown.Lines, $grown.States, $may, 1, 'FALSE')
Ok 'applying the same month again is idempotent' ([Rdv3Ledger]::SameLedger($grown.Lines, $grown.States, $again.Lines, $again.States)) ($again.Lines -join '|')
$preserveCfg = [Rdv3Config]::Load((Variant 'preserve-state' '"onSourceChange": "reset"' '"onSourceChange": "preserve"'))
$preserved = [Rdv3Ledger]::ApplyUpdate($preserveCfg.Data.UpdateJob, $april, $aprilStates, $may, 1, 'FALSE')
Ok 'merge uses the configured preserve rule rather than a built-in reset' (($preserved.States[1] -eq 'TRUE') -and ($preserved.ResetLines.Count -eq 0)) (($preserved.States -join ',') + ' reset=' + $preserved.ResetLines.Count)
$defaultKeep = [Rdv3Config]::Load((Variant 'defaultkeep' ', "targetOnly": "keep" }' ' }'))
Ok 'destination-only rows default to keep' ($defaultKeep.Data.UpdateJob.ApplyStep.TargetOnly -eq 'keep') $defaultKeep.Data.UpdateJob.ApplyStep.TargetOnly
$cfg.Data.UpdateJob.ApplyStep.TargetOnly = 'delete'
$synced = [Rdv3Ledger]::ApplyUpdate($cfg.Data.UpdateJob, $april, $aprilStates, $may, 1, 'FALSE')
Ok 'destination-only rows are deleted only when the step says delete' (($synced.Lines.Length -eq 2) -and ($synced.Deleted -eq 1)) (($synced.Lines -join '|') + ' deleted=' + $synced.Deleted)
$cfg.Data.UpdateJob.ApplyStep.TargetOnly = 'keep'

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
$m = ErrorOf { [Rdv3Table]::Read((Csv 'unicode-strict' "key1,value`r`n東京,a`r`n"), 'unicode-strict', $utf8, 'key1') }
Ok 'an omitted key rule still refuses a non-ASCII key' ($m -match 'unicode-strict.csv.*2.*ASCII') $m
$relaxedPath = Csv 'relaxed-keys' "key1,value`r`n青,first`r`n,blank`r`n東京,second`r`n青,later`r`n"
$relaxedTable = [Rdv3Table]::Read($relaxedPath, 'relaxed', $utf8, 'key1', $relaxedRules)
Ok 'relaxed CSV keys allow Unicode and variable character lengths' (($relaxedTable.Rows -eq 2) -and ($relaxedTable.KeyLen -eq 0) -and ($relaxedTable.Key(0) -eq '青') -and ($relaxedTable.Key(1) -eq '東京')) (($relaxedTable.Rows).ToString() + ' rows')
Ok 'relaxed CSV keys skip blanks and keep the first duplicate row' (($relaxedTable.Field(0, 1) -eq 'first') -and ($relaxedTable.Field(1, 1) -eq 'second') -and ($relaxedTable.SourceRow(1) -eq 4)) ($relaxedTable.Field(0, 1) + '/' + $relaxedTable.Field(1, 1))
$relaxedIndex = New-Object Rdv3Index $relaxedTable
$query = $utf8.GetBytes('東京'); $relaxedRows = $null
$relaxedCount = $relaxedIndex.FindBytes($query, 0, $query.Length, [ref]$relaxedRows)
Ok 'a relaxed table index decodes a variable-width lookup key' (($relaxedIndex.Keys -eq 2) -and ($relaxedCount -eq 1) -and ($relaxedRows[0] -eq 1)) ("keys/hits " + $relaxedIndex.Keys + '/' + $relaxedCount)
$sjis = [Text.Encoding]::GetEncoding(932)
$sj = Join-Path $work 'sjis.csv'
[IO.File]::WriteAllText($sj, "key1,名前`r`n000001,漢字`r`n", $sjis)
$ts = [Rdv3Table]::Read($sj, 'sjis', $sjis, 'key1')
Ok 'a Shift_JIS table reads with its encoding'   (($ts.Head[1] -eq '名前') -and ($ts.Field(0, 1) -eq '漢字')) ($ts.Head -join ',')
$bom = Join-Path $work 'bom.csv'
[IO.File]::WriteAllText($bom, "key1,f1`r`n000001,a`r`n", (New-Object Text.UTF8Encoding($true)))
$tb = [Rdv3Table]::Read($bom, 'bom', $utf8, 'key1')
Ok 'a UTF-8 BOM is skipped'                      (($tb.Head[0] -eq 'key1') -and ([Rdv3Table]::ReadHead($bom, $utf8)[0] -eq 'key1')) ($tb.Head -join ',')

$sourceBook = Join-Path $work 'source-table.xlsx'
[Rdv3Xlsx]::Write($sourceBook, [string[]]@('when','memo'), 'key1',
  [string[]]@("20240101`ta,b", "20240131`tz"), [string[]]@('K001','K002'), 'source')
$tx = [Rdv3Table]::Read($sourceBook, 'X', [Text.Encoding]::Unicode, 'key1')
Ok 'an xlsx source uses its first row as the table header' (($tx.Head -join ',') -eq 'key1,when,memo') ($tx.Head -join ',')
Ok 'an xlsx source keeps cells and ignores the CSV encoding' (($tx.Rows -eq 2) -and ($tx.Field(0, 0) -eq 'K001') -and ($tx.Field(0, 2) -eq 'a,b')) ($tx.Field(0, 0) + '/' + $tx.Field(0, 2))
$txIndex = New-Object Rdv3Index $tx
Ok 'an xlsx source is indexed by its declared unique key' (($txIndex.Find('K002').Count -eq 1) -and ($txIndex.Find('K002')[0] -eq 1)) 'xlsx key index'
$dupBook = Join-Path $work 'source-duplicate.xlsx'
[Rdv3Xlsx]::Write($dupBook, [string[]]@('value'), 'key1', [string[]]@('one','two'), [string[]]@('K001','K001'), 'duplicate')
$m = ErrorOf { New-Object Rdv3Index ([Rdv3Table]::Read($dupBook, 'X', $utf8, 'key1')) }
Ok 'an xlsx source refuses a duplicate key with both rows' ($m -match 'source-duplicate.xlsx.*K001.*2.*3') $m
$relaxedBook = Join-Path $work 'source-relaxed.xlsx'
[Rdv3Xlsx]::Write($relaxedBook, [string[]]@('value'), 'key1',
  [string[]]@('first','blank','second','later'), [string[]]@('青','','東京','青'), 'relaxed')
$relaxedXlsx = [Rdv3Table]::Read($relaxedBook, 'X', $utf8, 'key1', $relaxedRules)
$relaxedXlsxIndex = New-Object Rdv3Index $relaxedXlsx
Ok 'an xlsx source applies the same relaxed key rules' (($relaxedXlsx.Rows -eq 2) -and ($relaxedXlsxIndex.Keys -eq 2) -and ($relaxedXlsx.Field(0, 1) -eq 'first') -and ($relaxedXlsx.Field(1, 1) -eq 'second')) (($relaxedXlsx.Rows).ToString() + ' rows')

# The shipped definition declares data.types; bind those declarations to the
# real input headings and exercise their values and export comparisons.
$typedCfg = $cfg
$shippingData = Join-Path $Root 'dist\app-csharp\data'
$typedHeads = New-Object 'string[][]' $typedCfg.Data.Tables.Count
for ($i = 0; $i -lt $typedCfg.Data.Tables.Count; $i++) {
  $typedHeads[$i] = [Rdv3Table]::ReadHead((Join-Path $shippingData $typedCfg.Data.Tables[$i].File), $typedCfg.Data.Enc)
}
$typedCfg.Data.Bind($typedHeads)
Ok 'declared date and number types load beside labels' (($typedCfg.Data.TypeOf('B.b_date').Format -eq 'yyyyMMdd') -and ($typedCfg.Data.TypeOf('B.b_qty').Type -eq 'number')) $typedCfg.Data.Describe()
Ok 'an undeclared column remains text' ($null -eq $typedCfg.Data.TypeOf('B.key1')) 'a type was inferred'
$badColumnCfg = [Rdv3Config]::Load((Variant 'typed-missing-column' '"B.b_date": { "type": "date", "format": "yyyyMMdd" }' '"B.no_such_column": { "type": "date", "format": "yyyyMMdd" }'))
$m = ErrorOf { $badColumnCfg.Data.Bind($typedHeads) }
Ok 'a declared type still names a real input column' ($m -match 'tableB.csv.*no_such_column.*data.types') $m

$miniTyped = Join-Path $work 'typed-tableB.csv'
$typedRow = New-Object string[] $typedHeads[1].Length
$typedRow[[Array]::IndexOf($typedHeads[1], 'key1')] = '00000001'
$typedRow[[Array]::IndexOf($typedHeads[1], 'key2')] = '00000002'
$typedRow[[Array]::IndexOf($typedHeads[1], 'b_date')] = '20240229'
$typedRow[[Array]::IndexOf($typedHeads[1], 'b_qty')] = '12.5'
[IO.File]::WriteAllText($miniTyped, (($typedHeads[1] -join ',') + "`r`n" + ($typedRow -join ',') + "`r`n"), $utf8)
$typedTables = New-Object 'Rdv3Table[]' $typedCfg.Data.Tables.Count
$typedTables[1] = [Rdv3Table]::Read($miniTyped, 'B', $typedCfg.Data.Enc, 'key2')
$m = ErrorOf { $typedCfg.Data.ValidateTypes($typedTables) }
Ok 'declared date and number values validate before use' ($null -eq $m) $m
$typedRow[[Array]::IndexOf($typedHeads[1], 'b_date')] = '20240230'
[IO.File]::WriteAllText($miniTyped, (($typedHeads[1] -join ',') + "`r`n" + ($typedRow -join ',') + "`r`n"), $utf8)
$typedTables[1] = [Rdv3Table]::Read($miniTyped, 'B', $typedCfg.Data.Enc, 'key2')
$m = ErrorOf { $typedCfg.Data.ValidateTypes($typedTables) }
Ok 'a bad date stops with its column and actual value' ($m -match 'B.b_date.*20240230') $m
$typedRow[[Array]::IndexOf($typedHeads[1], 'b_date')] = '20240229'
$typedRow[[Array]::IndexOf($typedHeads[1], 'b_qty')] = 'twelve'
[IO.File]::WriteAllText($miniTyped, (($typedHeads[1] -join ',') + "`r`n" + ($typedRow -join ',') + "`r`n"), $utf8)
$typedTables[1] = [Rdv3Table]::Read($miniTyped, 'B', $typedCfg.Data.Enc, 'key2')
$m = ErrorOf { $typedCfg.Data.ValidateTypes($typedTables) }
Ok 'a bad number stops with its column and actual value' ($m -match 'B.b_qty.*twelve') $m

$filterValues = New-Object string[] $typedCfg.Data.Columns.Count
$filterValues[$typedCfg.Data.IndexOf('B.b_date')] = '20240131'
$filterValues[$typedCfg.Data.IndexOf('B.b_qty')] = '10'
$filterValues[$typedCfg.Data.IndexOf('A.a_name')] = 'Alpha Beta'
$dateFilter = New-Object Rdv3ExportFilter
$dateFilter.Field = 'B.b_date'; $dateFilter.Operator = 'range'; $dateFilter.First = '20240101'; $dateFilter.Last = '20240131'
Ok 'a date range includes both endpoints' ($dateFilter.Matches($typedCfg.Data, $filterValues, 'FALSE')) 'date endpoint missed'
$dateFilter.Last = '20240130'
Ok 'a date range excludes a value beyond its last day' (-not $dateFilter.Matches($typedCfg.Data, $filterValues, 'FALSE')) 'date outlier included'
$numberFilter = New-Object Rdv3ExportFilter
$numberFilter.Field = 'B.b_qty'; $numberFilter.Operator = 'range'; $numberFilter.First = '10'; $numberFilter.Last = '20'
Ok 'a number range includes both endpoints' ($numberFilter.Matches($typedCfg.Data, $filterValues, 'FALSE')) 'number endpoint missed'
$textFilter = New-Object Rdv3ExportFilter
$textFilter.Field = 'A.a_name'; $textFilter.First = 'Alpha'
$textFilter.Operator = 'contains'; $containsText = $textFilter.Matches($typedCfg.Data, $filterValues, 'FALSE')
$textFilter.Operator = 'startsWith'; $startsText = $textFilter.Matches($typedCfg.Data, $filterValues, 'FALSE')
$textFilter.Operator = 'equals'; $equalsText = $textFilter.Matches($typedCfg.Data, $filterValues, 'FALSE')
$textFilter.Operator = 'notContains'; $notContainsText = $textFilter.Matches($typedCfg.Data, $filterValues, 'FALSE')
Ok 'text filters provide all four literal comparisons' ($containsText -and $startsText -and -not $equalsText -and -not $notContainsText) 'text operators'
$andRequest = New-Object Rdv3ExportRequest
$dateFilter.Last = '20240131'
$textFilter.Operator = 'startsWith'
$andRequest.Filters.Add($dateFilter); $andRequest.Filters.Add($numberFilter); $andRequest.Filters.Add($textFilter)
$allMatch = $andRequest.Matches($typedCfg.Data, $filterValues, 'FALSE')
$filterValues[$typedCfg.Data.IndexOf('A.a_name')] = 'Beta Alpha'
Ok 'several export filters are combined with AND' ($allMatch -and -not $andRequest.Matches($typedCfg.Data, $filterValues, 'FALSE')) 'filters were not AND'

# ===========================================================================
Section 'every table-operation word is reachable through one typed pipeline'
$vocabDir = Join-Path $work 'vocabulary'
New-Item -ItemType Directory -Path $vocabDir | Out-Null
[IO.File]::WriteAllText((Join-Path $vocabDir 'x.csv'), "id,grp,qty,price,status`r`nX1,G1,2,5,OPEN`r`nX2,G1,3,4,CLOSED`r`nX3,G2,1,7,OPEN`r`n", $utf8)
[IO.File]::WriteAllText((Join-Path $vocabDir 'y.csv'), "id,grp,qty,price,status`r`nX1,G1,2,2,OPEN`r`nY2,G2,1,8,OPEN`r`n", $utf8)
[IO.File]::WriteAllText((Join-Path $vocabDir 'z.csv'), "id,note`r`nX1,N1`r`nZ1,NZ`r`n", $utf8)
$vocabJson = @'
{
  "tables": {
    "X": { "label": "X", "file": "x.csv", "key": "id" },
    "Y": { "label": "Y", "file": "y.csv", "key": "id" },
    "Z": { "label": "Z", "file": "z.csv", "key": "id" }
  },
  "labels": {
    "X.id": "X id", "Y.id": "Y id", "Z.id": "Z id", "X.grp": "group",
    "X.qty": "quantity", "X.price": "price", "X.status": "status", "Z.note": "note",
    "ledger": "ledger", "baseStacked": "base stacked", "baseRows": "base rows", "stacked": "stacked",
    "picked": "picked", "picked.group": "group", "picked.id": "id",
    "picked.qty": "quantity", "picked.price": "price", "picked.status": "status",
    "priced": "priced", "priced.amount": "amount", "openRows": "open rows",
    "groupOneRows": "group one rows", "openGroupOneRows": "open group one rows",
    "openOutsideGroupOneRows": "open outside group one rows",
    "revised": "revised", "dedup": "deduplicated", "totals": "totals",
    "totals.total": "total", "totals.count": "count", "ordered": "ordered",
    "innerRows": "inner rows", "leftRows": "left rows", "fullRows": "full rows",
    "rightOnly": "right-only rows", "sortedLedger": "sorted ledger",
    "uniqueLedger": "unique ledger", "ledgerOpen": "open ledger rows"
  },
  "jobs": [
    {
      "id": "base", "name": "base", "kind": "update",
      "inputs": [ { "table": "X" }, { "table": "Y" } ],
      "steps": [
        { "operation": "append", "target1": "X", "target2": "Y", "condition": "", "output": "baseStacked" },
        { "operation": "distinct", "target1": "baseStacked", "condition": "", "columns": ["X.id"], "output": "baseRows" },
        { "operation": "replace", "target1": "baseRows", "target2": "ledger", "keys": ["X.id", "X.id"], "condition": "", "output": "ledger" }
      ]
    },
    {
      "id": "alternate", "name": "alternate", "kind": "update",
      "inputs": [ { "table": "X" } ],
      "steps": [
        { "operation": "replace", "target1": "X", "target2": "ledger", "keys": ["X.id", "X.id"], "condition": "", "output": "ledger" }
      ]
    },
    {
      "id": "table-words", "name": "table words", "kind": "delete",
      "inputs": [ { "table": "X" }, { "table": "Y" } ],
      "steps": [
        { "operation": "append", "target1": "X", "target2": "Y", "condition": "", "output": "stacked" },
        { "operation": "select", "target1": "stacked", "condition": "", "columns": [
          { "column": "X.grp", "as": "group" }, { "column": "X.id", "as": "id" },
          { "column": "X.qty", "as": "qty" }, { "column": "X.price", "as": "price" },
          { "column": "X.status", "as": "status" }
        ], "output": "picked" },
        { "operation": "calculate", "target1": "picked", "condition": "", "column": "amount", "expression": "picked.qty * picked.price", "output": "priced" },
        { "operation": "extract", "target1": "priced", "condition": "", "where": { "column": "picked.status", "operator": "equals", "value": "OPEN" }, "output": "openRows" },
        { "operation": "extract", "target1": "priced", "condition": "", "where": { "column": "picked.group", "operator": "equals", "value": "G1" }, "output": "groupOneRows" },
        { "operation": "extract", "target1": "openRows", "target2": "groupOneRows", "condition": "both", "output": "openGroupOneRows" },
        { "operation": "extract", "target1": "openRows", "target2": "groupOneRows", "condition": "exclude", "output": "openOutsideGroupOneRows" },
        { "operation": "update", "target1": "priced", "target2": "openRows", "condition": "", "set": [
          { "column": "priced.amount", "expression": "priced.amount + 1" },
          { "column": "picked.status", "expression": "'ACTIVE'" }
        ], "output": "revised" },
        { "operation": "distinct", "target1": "revised", "condition": "", "columns": ["picked.id"], "output": "dedup" },
        { "operation": "aggregate", "target1": "dedup", "condition": "", "groupBy": ["picked.group"], "aggregates": [
          { "function": "sum", "column": "priced.amount", "as": "total" },
          { "function": "count", "as": "count" }
        ], "output": "totals" },
        { "operation": "sort", "target1": "totals", "condition": "", "orders": [
          { "column": "totals.total", "direction": "descending", "type": "number" }
        ], "output": "ordered" }
      ]
    },
    {
      "id": "join-words", "name": "join words", "kind": "delete",
      "inputs": [ { "table": "X" }, { "table": "Z" } ],
      "steps": [
        { "operation": "join", "target1": "X", "target2": "Z", "keys": ["X.id", "Z.id"], "condition": "match", "output": "innerRows" },
        { "operation": "join", "target1": "X", "target2": "Z", "keys": ["X.id", "Z.id"], "condition": "left", "output": "leftRows" },
        { "operation": "join", "target1": "X", "target2": "Z", "keys": ["X.id", "Z.id"], "condition": "full", "output": "fullRows" },
        { "operation": "extract", "target1": "fullRows", "condition": "", "where": { "column": "X.id", "operator": "empty" }, "output": "rightOnly" }
      ]
    },
    {
      "id": "direct-ledger-update", "name": "direct ledger update", "kind": "delete",
      "inputs": [],
      "steps": [
        { "operation": "sort", "target1": "ledger", "condition": "", "orders": [
          { "column": "X.id", "direction": "ascending", "type": "text" }
        ], "output": "sortedLedger" },
        { "operation": "distinct", "target1": "sortedLedger", "condition": "", "columns": ["X.id"], "output": "uniqueLedger" },
        { "operation": "extract", "target1": "uniqueLedger", "condition": "", "where": { "column": "X.status", "operator": "equals", "value": "OPEN" }, "output": "ledgerOpen" },
        { "operation": "update", "target1": "uniqueLedger", "target2": "ledgerOpen", "condition": "", "set": [
          { "column": "X.price", "expression": "X.price + 1" }
        ], "output": "ledger" }
      ]
    }
  ],
  "ledger": {
    "identity": "X.id",
    "search": { "columns": ["X.id"], "match": "exact" },
    "columns": {
      "application": [ { "name": "workState", "onSourceChange": "reset" } ],
      "source": ["X.id", "X.grp", "X.qty", "X.price", "X.status", "Z.note"]
    }
  }
}
'@
$vocab = [Rdv3Data]::Read([Rdv3Json]::Parse($vocabJson))
$vocab.Bind(@(
  [string[]]@('id','grp','qty','price','status'),
  [string[]]@('id','grp','qty','price','status'),
  [string[]]@('id','note')))
$genericMerge = [Rdv3Ledger]::BuildFromCsv($vocab, $vocabDir)
Ok 'an update job can execute append and distinct before its ledger write' ((-not $vocab.UpdateJob.FastJoinPlan) -and ($genericMerge.Rows -eq 4) -and ($genericMerge.Lines[3] -match '^Y2')) ($genericMerge.Lines -join '|')
$alternateMerge = [Rdv3Ledger]::BuildFromCsv($vocab, $vocab.JobOf('alternate'), $vocabDir)
Ok 'the requested update job runs instead of silently returning to the first one' (($alternateMerge.Rows -eq 3) -and (($alternateMerge.Lines -join '|') -notmatch 'Y2')) ($alternateMerge.Lines -join '|')
$wordRun = [Rdv3Process]::Run($vocab, $vocab.JobOf('table-words'), $vocabDir, [string[]]@(), [string[]]@(), 'FALSE')
Ok 'append is executable, not a declared-only word' (($wordRun.ValueOf('stacked').Count -eq 5) -and ($wordRun.ValueOf('stacked').Lines[3] -match '^X1')) ($wordRun.ValueOf('stacked').Lines -join '|')
Ok 'select chooses, reorders, and renames columns' (($wordRun.ValueOf('picked').Columns -join ',') -eq 'picked.group,picked.id,picked.qty,picked.price,picked.status') ($wordRun.ValueOf('picked').Columns -join ',')
Ok 'calculate evaluates the configured arithmetic expression' (($wordRun.ValueOf('priced').Lines[0] -eq "G1`tX1`t2`t5`tOPEN`t10") -and ($wordRun.ValueOf('priced').Lines[1] -match "`t12$")) ($wordRun.ValueOf('priced').Lines -join '|')
Ok 'extract filters a table column into a bounded row set' ($wordRun.ValueOf('openRows').Count -eq 4) $wordRun.ValueOf('openRows').Count
Ok 'row-set both keeps the intersection' ($wordRun.ValueOf('openGroupOneRows').Count -eq 2) $wordRun.ValueOf('openGroupOneRows').Count
Ok 'row-set exclude removes the second set from the first' ($wordRun.ValueOf('openOutsideGroupOneRows').Count -eq 2) $wordRun.ValueOf('openOutsideGroupOneRows').Count
Ok 'update rewrites only the selected rows and is reachable in a delete-kind job' (($wordRun.ValueOf('revised').Lines[0] -eq "G1`tX1`t2`t5`tACTIVE`t11") -and ($wordRun.ValueOf('revised').Lines[1] -eq "G1`tX2`t3`t4`tCLOSED`t12")) ($wordRun.ValueOf('revised').Lines -join '|')
Ok 'distinct keeps the first row for each configured column tuple' ($wordRun.ValueOf('dedup').Count -eq 4) ($wordRun.ValueOf('dedup').Lines -join '|')
Ok 'aggregate provides sum and count by configured groups' (($wordRun.ValueOf('totals').Lines -join '|') -eq "G1`t23`t2|G2`t17`t2") ($wordRun.ValueOf('totals').Lines -join '|')
Ok 'sort orders by configured numeric descending order' (($wordRun.Lines[0] -eq "G1`t23`t2") -and ($wordRun.Lines[1] -eq "G2`t17`t2")) ($wordRun.Lines -join '|')
$joinRun = [Rdv3Process]::Run($vocab, $vocab.JobOf('join-words'), $vocabDir, [string[]]@(), [string[]]@(), 'FALSE')
Ok 'inner join keeps only paired rows' ($joinRun.ValueOf('innerRows').Count -eq 1) ($joinRun.ValueOf('innerRows').Lines -join '|')
Ok 'left outer join keeps every left row' (($joinRun.ValueOf('leftRows').Count -eq 3) -and ($joinRun.ValueOf('leftRows').Lines[1] -match "`t`t$")) ($joinRun.ValueOf('leftRows').Lines -join '|')
Ok 'full outer join also keeps right-only rows' (($joinRun.ValueOf('fullRows').Count -eq 4) -and ($joinRun.ValueOf('fullRows').Lines[3] -match "^`t`t`t`t`tZ1`tNZ$")) ($joinRun.ValueOf('fullRows').Lines -join '|')
Ok 'outer-join blanks can be filtered like ordinary empty fields' ($joinRun.ValueOf('rightOnly').Count -eq 1) $joinRun.ValueOf('rightOnly').Count
$directInput = [string[]]@("X9`tG9`t1`t4`tOPEN`t")
$directStates = [string[]]@('TRUE')
$directRun = [Rdv3Process]::Run($vocab, $vocab.JobOf('direct-ledger-update'), $vocabDir, $directInput, $directStates, 'FALSE')
Ok 'a direct source-column update obeys the configured reset rule' (($directRun.Lines[0] -eq "X9`tG9`t1`t5`tOPEN`t") -and ($directRun.States[0] -eq 'FALSE') -and ($directRun.Update.Updated -eq 1) -and ($directRun.Update.ResetLines.Count -eq 1)) (($directRun.Lines -join '|') + ' / ' + ($directRun.States -join ','))
$vocabPreserve = [Rdv3Data]::Read([Rdv3Json]::Parse($vocabJson.Replace('"onSourceChange": "reset"', '"onSourceChange": "preserve"')))
$vocabPreserve.Bind(@(
  [string[]]@('id','grp','qty','price','status'),
  [string[]]@('id','grp','qty','price','status'),
  [string[]]@('id','note')))
$directPreserved = [Rdv3Process]::Run($vocabPreserve, $vocabPreserve.JobOf('direct-ledger-update'), $vocabDir, $directInput, $directStates, 'FALSE')
Ok 'a direct source-column update obeys the configured preserve rule' (($directPreserved.States[0] -eq 'TRUE') -and ($directPreserved.Update.ResetLines.Count -eq 0)) (($directPreserved.States -join ',') + ' reset=' + $directPreserved.Update.ResetLines.Count)

# ===========================================================================
Section 'string functions create ordinary columns for later table operations'
$stringDir = Join-Path $work 'string-functions'
New-Item -ItemType Directory -Path $stringDir | Out-Null
[IO.File]::WriteAllText((Join-Path $stringDir 'source.csv'), "id,raw,status`r`nS1,ORD-AB12-TAIL,OPEN`r`nS2,ORD-CD34-TAIL,OPEN`r`nS3,ORD-EF56-TAIL,OPEN`r`n", $utf8)
[IO.File]::WriteAllText((Join-Path $stringDir 'lookup.csv'), "key,note`r`nAB12,first`r`nEF56,last`r`n", $utf8)
$stringJson = @'
{
  "tables": {
    "A": { "label": "source", "file": "source.csv", "key": "id" },
    "B": { "label": "lookup", "file": "lookup.csv", "key": "key" }
  },
  "labels": {
    "A.id": "source id", "A.raw": "source text", "A.status": "status",
    "B.key": "lookup key", "B.note": "lookup note", "ledger": "ledger",
    "regexed": "regex output", "regexed.regexKey": "regex key",
    "split": "split output", "split.splitKey": "split key",
    "derived": "derived output", "derived.substringKey": "substring key",
    "joined": "joined", "matchedRows": "matched rows", "updated": "updated",
    "matchedAfterUpdate": "matched after update", "remaining": "remaining"
  },
  "jobs": [
    {
      "id": "base", "name": "base", "kind": "update", "inputs": [ { "table": "A" } ],
      "steps": [
        { "operation": "replace", "target1": "A", "target2": "ledger",
          "keys": ["A.id", "A.id"], "condition": "", "output": "ledger" }
      ]
    },
    {
      "id": "strings", "name": "strings", "kind": "delete",
      "inputs": [ { "table": "A" }, { "table": "B" } ],
      "steps": [
        { "operation": "calculate", "target1": "A", "condition": "", "column": "regexKey",
          "expression": "regexExtract(A.raw, '[A-Z]{2}[0-9]{2}')", "output": "regexed" },
        { "operation": "calculate", "target1": "regexed", "condition": "", "column": "splitKey",
          "expression": "splitPart(A.raw, '-', 1)", "output": "split" },
        { "operation": "calculate", "target1": "split", "condition": "", "column": "substringKey",
          "expression": "substring(split.splitKey, 0, 4)", "output": "derived" },
        { "operation": "join", "target1": "derived", "target2": "B",
          "keys": ["derived.substringKey", "B.key"], "condition": "match", "output": "joined" },
        { "operation": "extract", "target1": "derived", "target2": "B",
          "keys": ["derived.substringKey", "B.key"], "condition": "match", "output": "matchedRows" },
        { "operation": "update", "target1": "derived", "target2": "matchedRows", "condition": "",
          "set": [ { "column": "A.status", "expression": "substring('MATCHED', 0, 7)" } ], "output": "updated" },
        { "operation": "extract", "target1": "updated", "target2": "B",
          "keys": ["derived.substringKey", "B.key"], "condition": "match", "output": "matchedAfterUpdate" },
        { "operation": "delete", "target1": "updated", "target2": "matchedAfterUpdate",
          "condition": "", "output": "remaining" }
      ]
    }
  ],
  "ledger": {
    "identity": "A.id", "search": { "columns": ["A.id"], "match": "exact" },
    "columns": { "application": [ { "name": "workState", "onSourceChange": "reset" } ],
      "source": ["A.id", "A.raw", "A.status"] }
  }
}
'@
$stringHeads = @([string[]]@('id','raw','status'), [string[]]@('key','note'))
$stringData = [Rdv3Data]::Read([Rdv3Json]::Parse($stringJson))
$stringData.Bind($stringHeads)
$stringRun = [Rdv3Process]::Run($stringData, $stringData.JobOf('strings'), $stringDir, [string[]]@(), [string[]]@(), 'FALSE')
$derived = $stringRun.ValueOf('derived')
Ok 'regexExtract, splitPart, and substring produce their configured values' (($derived.Lines[0] -eq "S1`tORD-AB12-TAIL`tOPEN`tAB12`tAB12`tAB12") -and ($derived.Lines[2] -match "`tEF56`tEF56`tEF56$")) ($derived.Lines -join '|')
Ok 'a generated value is retained as an ordinary named column' (($derived.Columns[$derived.Columns.Count - 1] -eq 'derived.substringKey') -and ($derived.Count -eq 3)) ($derived.Columns -join ',')
Ok 'a generated column is usable as a later join key' ($stringRun.ValueOf('joined').Count -eq 2) $stringRun.ValueOf('joined').Count
Ok 'a generated column is usable for table-to-table matching' ($stringRun.ValueOf('matchedRows').Count -eq 2) $stringRun.ValueOf('matchedRows').Count
Ok 'rows selected by a generated key can be updated with a string function' (($stringRun.ValueOf('updated').Lines[0] -match "^S1`t.*`tMATCHED`t") -and ($stringRun.ValueOf('updated').Lines[1] -match "^S2`t.*`tOPEN`t") -and ($stringRun.ValueOf('updated').Lines[2] -match "^S3`t.*`tMATCHED`t")) ($stringRun.ValueOf('updated').Lines -join '|')
Ok 'rows selected again by the generated key can be deleted' (($stringRun.ValueOf('remaining').Count -eq 1) -and ($stringRun.ValueOf('remaining').Lines[0] -match '^S2')) ($stringRun.ValueOf('remaining').Lines -join '|')
$displayFunction = [Rdv3Process]::DisplayExpression($stringData, "splitPart(A.raw, '-', 1)")
Ok 'function display keeps the function and substitutes the configured column label' ($displayFunction -eq "splitPart(source text, '-', 1)") $displayFunction

function ExpressionFailure([string] $expression) {
  $candidate = $stringJson.Replace("regexExtract(A.raw, '[A-Z]{2}[0-9]{2}')", $expression)
  try {
    $data = [Rdv3Data]::Read([Rdv3Json]::Parse($candidate))
    $data.Bind($stringHeads)
    [void][Rdv3Process]::Run($data, $data.JobOf('strings'), $stringDir, [string[]]@(), [string[]]@(), 'FALSE')
    return ''
  } catch { return $_.Exception.Message }
}
$m = ExpressionFailure "notAFunction(A.raw)"
Ok 'an unknown expression function is refused' ($m -match 'unknown.*function') $m
$m = ExpressionFailure "regexExtract(A.raw)"
Ok 'a string function requires its exact argument count' ($m -match 'regexExtract expects 2 arguments') $m
$m = ExpressionFailure "regexExtract(A.raw, A.raw)"
Ok 'a regular expression argument must be quoted text' ($m -match 'pattern must be quoted text') $m
$m = ExpressionFailure "regexExtract(A.raw, '[')"
Ok 'an unusable extraction pattern is refused before execution' ($m -match 'not a usable regular expression') $m
$m = ExpressionFailure "splitPart(A.raw, '-', 1.5)"
Ok 'a split position must be a whole number' ($m -match 'position must be a non-negative whole number') $m
$m = ExpressionFailure "substring(A.raw, 0, 0)"
Ok 'a substring length must be positive' ($m -match 'length must be a positive whole number') $m
$m = ExpressionFailure "regexExtract('', '[A-Z]+')"
Ok 'an empty source value stops string extraction' ($m -match 'regexExtract received empty text') $m
$m = ExpressionFailure "regexExtract(A.raw, '^Z[0-9]+$')"
Ok 'a regular expression mismatch stops instead of making an empty key' ($m -match 'regexExtract found no match') $m
$m = ExpressionFailure "splitPart(A.raw, '-', 9)"
Ok 'a split position outside the result is refused' ($m -match 'splitPart position is outside') $m
$m = ExpressionFailure "splitPart('A--B', '-', 1)"
Ok 'an empty split element stops instead of making an empty key' ($m -match 'splitPart produced empty text') $m
$m = ExpressionFailure "substring(A.raw, 99, 1)"
Ok 'a substring range outside the source is refused' ($m -match 'substring range is outside') $m

# ===========================================================================
Section 'the merge follows the definition: spine, joins, ledger columns'
$dd = Join-Path $work 'tiny'
New-Item -ItemType Directory -Path $dd | Out-Null
[IO.File]::WriteAllText((Join-Path $dd 'a.csv'), "key1,a_x`r`n0001,AX1`r`n0002,AX2`r`n", $utf8)
[IO.File]::WriteAllText((Join-Path $dd 'b.csv'), "key1,key2,b_x`r`n0001,K1,B1`r`n0001,K2,B2`r`n0009,K3,B3`r`n", $utf8)
[IO.File]::WriteAllText((Join-Path $dd 'c.csv'), "key2,c_x,c_y`r`nK1,C1,Y1`r`nK3,C3,Y3`r`n", $utf8)
$defJson = '{ "tables": { "A": { "label": "A", "file": "a.csv", "key": "key1" }, "B": { "label": "B", "file": "b.csv", "key": "key2" }, "C": { "label": "C", "file": "c.csv", "key": "key2" } },' +
           '  "labels": { "A.key1": "A key1", "B.key1": "key1", "B.key2": "key2", "C.key2": "C key2", "ledger": "ledger", "previousLedger": "previous", "m1": "m1", "m2": "m2" },' +
           '  "jobs": [ { "id": "replace", "name": "replace", "kind": "update", "inputs": [ { "table": "A" }, { "table": "B" }, { "table": "C" } ], "steps": [' +
           '    { "operation": "join", "target1": "B", "target2": "A", "keys": ["B.key1", "A.key1"], "condition": "left", "output": "m1" },' +
           '    { "operation": "join", "target1": "m1", "target2": "C", "keys": ["B.key2", "C.key2"], "condition": "left", "output": "m2" },' +
           '    { "operation": "replace", "target1": "m2", "target2": "ledger", "keys": ["B.key2", "B.key2"], "condition": "", "output": "ledger" } ] } ],' +
           '  "ledger": { "identity": "B.key2", "search": { "columns": ["B.key1"], "match": "exact" }, "columns": { "application": [ { "name": "workState", "onSourceChange": "reset" } ], "source": [ "B.key2", "C.c_y", "A.a_x", "B.b_x", "B.key1" ] } } }'
$def = [Rdv3Data]::Read([Rdv3Json]::Parse($defJson))
$mr = [Rdv3Ledger]::BuildFromCsv($def, $dd)
Ok 'one ledger row per spine row'        ($mr.Rows -eq 3) ("rows " + $mr.Rows)
Ok 'the columns come in ledger order'    ($mr.Lines[0] -eq "K1`tY1`tAX1`tB1`t0001") $mr.Lines[0]
Ok 'a row whose join finds nothing is blank there' ($mr.Lines[1] -eq "K2`t`tAX1`tB2`t0001") $mr.Lines[1]
Ok 'a spine key nobody has leaves A blank' ($mr.Lines[2] -eq "K3`tY3`t`tB3`t0009") $mr.Lines[2]
Ok 'the matched counts say so'           (($mr.Matched[0] -eq 2) -and ($mr.Matched[1] -eq 2)) ($mr.Matched -join ',')
Ok 'the xlsx header is the CSV names'    (($mr.Head -join ',') -eq 'key2,c_y,a_x,b_x,key1') ($mr.Head -join ',')
Ok 'identity and search columns follow'  (($def.IdentityCol -eq 0) -and ($def.SearchCols[0] -eq 4)) ("{0}/{1}" -f $def.IdentityCol, $def.SearchCols[0])
$inputBook = Join-Path $dd 'b.xlsx'
[Rdv3Xlsx]::Write($inputBook, [string[]]@('key2','b_x'), 'key1',
  [string[]]@("K1`tB1", "K2`tB2", "K3`tB3"), [string[]]@('0001','0001','0009'), 'table-b')
$xlsxDef = [Rdv3Data]::Read([Rdv3Json]::Parse($defJson.Replace('"b.csv"', '"b.xlsx"')))
$xlsxMerge = [Rdv3Ledger]::BuildFromCsv($xlsxDef, $dd)
Ok 'an xlsx table follows the same join and ledger path as CSV' (($xlsxMerge.Lines -join '|') -eq ($mr.Lines -join '|')) ($xlsxMerge.Lines -join '|')
$relaxedJoinDir = Join-Path $work 'relaxed-join'
New-Item -ItemType Directory -Path $relaxedJoinDir | Out-Null
[IO.File]::WriteAllText((Join-Path $relaxedJoinDir 'a.csv'), "key1,a_x`r`n青,AX1`r`n東京,AX2`r`n青,IGNORED`r`n,BLANK`r`n", $utf8)
[IO.File]::WriteAllText((Join-Path $relaxedJoinDir 'b.csv'), "key1,key2,b_x`r`n青,K1,B1`r`n東京,K2,B2`r`n", $utf8)
[IO.File]::WriteAllText((Join-Path $relaxedJoinDir 'c.csv'), "key2,c_x,c_y`r`nK1,C1,Y1`r`nK2,C2,Y2`r`n", $utf8)
$relaxedDefJson = $defJson.Replace('"A": { "label": "A", "file": "a.csv", "key": "key1" }',
  '"A": { "label": "A", "file": "a.csv", "key": "key1", "keyValidation": { "characters": "unicode", "length": "variable", "duplicates": "distinct", "empty": "skip" } }')
$relaxedDef = [Rdv3Data]::Read([Rdv3Json]::Parse($relaxedDefJson))
$relaxedMerge = [Rdv3Ledger]::BuildFromCsv($relaxedDef, $relaxedJoinDir)
Ok 'relaxed table keys flow through the ordinary configured join' (($relaxedMerge.Rows -eq 2) -and ($relaxedMerge.Keys[0] -eq 2) -and ($relaxedMerge.Lines[0] -eq "K1`tY1`tAX1`tB1`t青") -and ($relaxedMerge.Lines[1] -eq "K2`tY2`tAX2`tB2`t東京")) ($relaxedMerge.Lines -join '|')
$replaceOld = [string[]]@("OLD`told", $mr.Lines[0])
$replaceStates = [string[]]@('TRUE', 'TRUE')
$replaced = [Rdv3Ledger]::ApplyUpdate($def.UpdateJob, $replaceOld, $replaceStates, $mr.Lines, 0, 'FALSE')
Ok 'replace remains available and follows source order' (($replaced.Lines.Length -eq 3) -and ($replaced.Lines[0] -eq $mr.Lines[0]) -and ($replaced.Deleted -eq 1)) (($replaced.Lines -join '|') + ' deleted=' + $replaced.Deleted)
Ok 'replace also preserves an unchanged application-owned value' ($replaced.States[0] -eq 'TRUE') ($replaced.States -join ',')
$m = ErrorOf { [Rdv3Data]::Read([Rdv3Json]::Parse($defJson.Replace('"source": [ "B.key2",', '"source": ['))) }
Ok 'a ledger without the identity is refused' ($m -match 'B.key2') $m
$def2 = [Rdv3Data]::Read([Rdv3Json]::Parse($defJson.Replace('"C.c_y"', '"C.c_z"')))
$m = ErrorOf { [Rdv3Ledger]::BuildFromCsv($def2, $dd) }
Ok 'a column the CSV lacks stops the merge, naming the file' ($m -match 'c.csv.*c_z') $m
[IO.File]::WriteAllText((Join-Path $dd 'a.csv'), "key1,a_x`r`n0001,AX1`r`n0001,AX2`r`n", $utf8)
$m = ErrorOf { [Rdv3Ledger]::BuildFromCsv($def, $dd) }
Ok 'a duplicate key in a joined table stops the merge' ($m -match 'a.csv.*0001') $m

# The same physical xlsx grows from an April-only source to April + May.
$replaceStep = '{ "operation": "replace", "target1": "m2", "target2": "ledger", "keys": ["B.key2", "B.key2"], "condition": "", "output": "ledger" }'
$mergeStep = '{ "operation": "merge", "target1": "m2", "target2": "ledger", "keys": ["B.key2", "B.key2"], "condition": "", "output": "ledger", "sourceOnly": "add", "both": "update", "targetOnly": "keep" }'
$mergeDef = [Rdv3Data]::Read([Rdv3Json]::Parse($defJson.Replace($replaceStep, $mergeStep)))
$monthPath = Join-Path $work 'month-ledger.xlsx'
$aprilLines = [string[]]@("K1`tY1`tAX1`tB1`t0001", "K2`t`tAX1`tB2`t0001")
[Rdv3Xlsx]::Write($monthPath, $mr.Head, 'state', $aprilLines, [string[]]@('TRUE','TRUE'), 'april')
$monthLines = $null; $monthStates = $null
[Rdv3Xlsx]::Read($monthPath, $mr.Head, 'state', [ref]$monthLines, [ref]$monthStates)
$mayLines = [string[]]@("K2`t`tAX1`tB2-NEW`t0001", "K3`tY3`t`tB3`t0009")
$monthResult = [Rdv3Ledger]::ApplyUpdate($mergeDef.UpdateJob, $monthLines, $monthStates, $mayLines, 0, 'FALSE')
[Rdv3Xlsx]::Write($monthPath, $mr.Head, 'state', $monthResult.Lines, $monthResult.States, 'may')
$grownLines = $null; $grownStates = $null
[Rdv3Xlsx]::Read($monthPath, $mr.Head, 'state', [ref]$grownLines, [ref]$grownStates)
Ok 'one xlsx grows from April to April plus May' (($grownLines.Length -eq 3) -and ($grownLines[0] -eq $aprilLines[0]) -and ($grownLines[1] -eq $mayLines[0]) -and ($grownLines[2] -eq $mayLines[1])) ($grownLines -join '|')
Ok 'the same-file merge keeps old marks and resets changed content' (($grownStates -join ',') -eq 'TRUE,FALSE,FALSE') ($grownStates -join ',')

# ===========================================================================
Section 'search uses every configured column and either exact or contains matching'
$searchLines = [string[]]@("AA100`tZZ900", "BB200`tALPHA", "ALPHA`tALPHA", "CC300`tOMEGA")
$ctor = [Rdv3Index].GetConstructor([Type[]]@([string[]], [int[]], [string]))
$exact = $ctor.Invoke([object[]]@($searchLines, [int[]]@(0,1), 'exact'))
$hits = $exact.Find('ALPHA')
Ok 'exact search finds values in both configured columns' (($hits.Count -eq 2) -and ($hits[0] -eq 1) -and ($hits[1] -eq 2)) ($hits -join ',')
Ok 'one row matching two columns appears only once' (($hits | Where-Object { $_ -eq 2 }).Count -eq 1) ($hits -join ',')
$contains = $ctor.Invoke([object[]]@($searchLines, [int[]]@(0,1), 'contains'))
$hits = $contains.Find('PHA')
Ok 'contains search scans the configured columns' (($hits.Count -eq 2) -and ($hits[0] -eq 1) -and ($hits[1] -eq 2)) ($hits -join ',')
Ok 'contains search returns no row when no column matches' ($null -eq $contains.Find('NOT-THERE')) 'no hits'

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

# ===========================================================================
Section 'delete runs only through row sets selected by configured ledger columns'
$deleteDir = Join-Path $work 'delete-inputs'
New-Item -ItemType Directory -Path $deleteDir | Out-Null
[IO.File]::WriteAllText((Join-Path $deleteDir 'delete.csv'), "key2,b_ref`r`nDEL00001,REF00001`r`nKEEP0002,REF00002`r`n", $utf8)
function ShippedLine([string] $key1, [string] $key2, [string] $ref) {
  $v = New-Object string[] $cfg.Data.Columns.Count
  $v[$cfg.Data.IndexOf('B.key1')] = $key1
  $v[$cfg.Data.IndexOf('B.key2')] = $key2
  $v[$cfg.Data.IndexOf('B.b_ref')] = $ref
  return ($v -join "`t")
}
$deleteLines = [string[]]@(
  (ShippedLine 'KEY00001' 'DEL00001' 'REF00001'),
  (ShippedLine 'KEY00002' 'KEEP0002' 'REF00002'),
  (ShippedLine 'KEY00003' 'KEEP0003' 'REF00003'))
$deleteStates = [string[]]@('TRUE', 'FALSE', 'TRUE')
$deleted = [Rdv3Ledger]::ApplyDelete($cfg.Data, $cfg.Data.JobOf('delete-listed-records'), $deleteDir, $deleteLines, $deleteStates, 'FALSE')
Ok 'the shipped paired columns are intersected before deletion' (($deleted.Deleted -eq 2) -and ($deleted.Lines.Length -eq 1)) (($deleted.Lines.Length).ToString() + ' rows')
Ok 'unselected content and application state stay together' (($deleted.Lines[0] -eq $deleteLines[2]) -and ($deleted.States[0] -eq 'TRUE')) ($deleted.States -join ',')
[IO.File]::WriteAllText((Join-Path $deleteDir 'conditions.csv'), "key2,b_ref`r`nDEL00001,NONE0001`r`nNONE0002,REF00002`r`n", $utf8)
$sameFileText = $shipped.Replace('"file": "delete.csv"', '"file": "conditions.csv"')
$sameFileCfg = [Rdv3Config]::Load((WriteText 'same-file-columns.json' $sameFileText))
$sameFileDeleted = [Rdv3Ledger]::ApplyDelete($sameFileCfg.Data, $sameFileCfg.Data.JobOf('delete-listed-records'), $deleteDir, $deleteLines, $deleteStates, 'FALSE')
Ok 'both keeps rows selected by only one of the two columns' (($sameFileDeleted.Deleted -eq 0) -and ($sameFileDeleted.Lines.Length -eq 3)) (($sameFileDeleted.Lines.Length).ToString() + ' rows')
[IO.File]::WriteAllText((Join-Path $deleteDir 'relaxed-values.csv'), "key2,b_ref`r`n東京,NONE0001`r`n,NONE0002`r`n東京,NONE0003`r`n大阪,NONE0004`r`n", $utf8)
$relaxedInput = '"file": "relaxed-values.csv", "column": "key2", "key": "B.key2", "keyValidation": { "characters": "unicode", "length": "variable", "duplicates": "distinct", "empty": "skip" }'
$relaxedInputText = $shipped.Replace('"file": "delete.csv", "column": "key2", "key": "B.key2"', $relaxedInput)
$relaxedInputText = $relaxedInputText.Replace('"file": "delete.csv", "column": "b_ref"', '"file": "relaxed-values.csv", "column": "b_ref"')
$relaxedInputCfg = [Rdv3Config]::Load((WriteText 'relaxed-input.json' $relaxedInputText))
$relaxedInputLines = [string[]]@(
  (ShippedLine 'KEY00001' '東京' 'NONE0001'),
  (ShippedLine 'KEY00002' '大阪' 'NONE0002'),
  (ShippedLine 'KEY00003' '京都' 'NONE0003'))
$relaxedInputDeleted = [Rdv3Ledger]::ApplyDelete($relaxedInputCfg.Data, $relaxedInputCfg.Data.JobOf('delete-listed-records'), $deleteDir, $relaxedInputLines, [string[]]@('TRUE','FALSE','TRUE'), 'FALSE')
Ok 'a relaxed condition-value input is distinct, blank-free, and usable by the generic pipeline' (($relaxedInputDeleted.Deleted -eq 2) -and ($relaxedInputDeleted.Lines.Length -eq 1) -and ($relaxedInputDeleted.Lines[0] -eq $relaxedInputLines[2])) (($relaxedInputDeleted.Lines.Length).ToString() + ' rows')

# ===========================================================================
Section 'local pending changes stay local, survive restart, and require the same record content'
$pendingPath = Join-Path $work 'pending.dat'
$pending = New-Object Rdv3PendingStore $pendingPath
$pending.Set('K001', 'TRUE', "K001`talpha", 'FALSE')
Ok 'one local change is stored' (($pending.Count -eq 1) -and (Test-Path -LiteralPath $pendingPath)) ("count " + $pending.Count)
$pending = New-Object Rdv3PendingStore $pendingPath
Ok 'the local change survives reopening the store' ($pending.Count -eq 1) ("count " + $pending.Count)
$overlay = $pending.Overlay([string[]]@("K001`talpha"), [string[]]@('FALSE'), 0)
Ok 'the local state overlays identical shared content' ($overlay[0] -eq 'TRUE') $overlay[0]
$overlay = $pending.Overlay([string[]]@("K001`tchanged"), [string[]]@('FALSE'), 0)
Ok 'changed shared content is not marked from the old pending entry' ($overlay[0] -eq 'FALSE') $overlay[0]
$pending.Set('K002', 'TRUE', "K002`tbeta", 'FALSE')
$apply = $pending.PrepareSend([string[]]@("K001`talpha"), [string[]]@('FALSE'), 0, 'FALSE')
Ok 'send resolves only the matching identity and content' (($apply.Resolved.Count -eq 1) -and ($apply.States[0] -eq 'TRUE')) ("resolved " + $apply.Resolved.Count)
Ok 'a missing identity is reported and retained' (($apply.Unmatched.Count -eq 1) -and ($apply.Unmatched[0].Identity -eq 'K002') -and ($apply.Unmatched[0].Reason -eq 'missing')) ("unmatched " + $apply.Unmatched.Count)
$pending.Remove($apply.Resolved)
Ok 'only resolved changes leave the local store' (($pending.Count -eq 1) -and ($pending.Snapshot()[0].Identity -eq 'K002')) ("count " + $pending.Count)
$pending.Set('K002', 'FALSE', "K002`tbeta", 'FALSE')
Ok 'returning to the shared value removes the pending change' ($pending.Count -eq 0) ("count " + $pending.Count)
$pending.Set('K003', 'FALSE', "K003`tgamma", 'TRUE')
$apply = $pending.PrepareSend([string[]]@("K003`tgamma"), [string[]]@('TRUE'), 0, 'FALSE')
Ok 'send carries a return to the initial state too' (($apply.States[0] -eq 'FALSE') -and ($apply.FromInitial -eq 0) -and ($apply.ToInitial -eq 1)) ("done " + $apply.FromInitial + ', todo ' + $apply.ToInitial)
$pending.Remove($apply.Resolved)
$pending.Set('K004', 'TRUE', "K004`told", 'FALSE')
$apply = $pending.PrepareSend([string[]]@("K004`tnew"), [string[]]@('FALSE'), 0, 'FALSE')
Ok 'changed content is reported rather than overwritten' (($apply.Resolved.Count -eq 0) -and ($apply.Unmatched.Count -eq 1) -and ($apply.Unmatched[0].Reason -eq 'changed')) $apply.Unmatched[0].Reason

# ===========================================================================
Section 'shared writes use one lock and publish a one-line version marker'
$sharedLedger = Join-Path $work 'shared.xlsx'
$shared1 = New-Object Rdv3SharedFiles $sharedLedger, 'HOST1', 'alice', 'writer-1'
$shared2 = New-Object Rdv3SharedFiles $sharedLedger, 'HOST2', 'bob', 'writer-2'
$owner = $null
$lock1 = $shared1.TryAcquire([ref]$owner)
Ok 'the first writer acquires CreateNew lock' (($null -ne $lock1) -and (Test-Path -LiteralPath ($sharedLedger + '.lock'))) 'first lock'
$owner = $null
$lock2 = $shared2.TryAcquire([ref]$owner)
Ok 'a second writer waits and can name the owner' (($null -eq $lock2) -and ($owner.User -eq 'alice') -and ($owner.Host -eq 'HOST1')) ($owner.User + '@' + $owner.Host)
$lock1.Release()
$owner = $null
$lock2 = $shared2.TryAcquire([ref]$owner)
Ok 'the second writer acquires after release' ($null -ne $lock2) 'second lock'
$mark1 = $shared2.WriteMarker('send', 2, 1, 1)
$mark2 = $shared2.WriteMarker('update', 3, 0, 0)
$readMark = $shared1.ReadMarker()
Ok 'marker versions advance under the same lock' (($mark1.Version -eq 1) -and ($mark2.Version -eq 2) -and ($readMark.Version -eq 2)) ("versions " + $mark1.Version + ',' + $mark2.Version)
Ok 'the marker carries writer, rows, and operation' (($readMark.User -eq 'bob') -and ($readMark.Host -eq 'HOST2') -and ($readMark.Rows -eq 3) -and ($readMark.Kind -eq 'update')) ($readMark.User + ' ' + $readMark.Kind)
Ok 'the marker is a single small line' (([IO.File]::ReadAllLines($sharedLedger + '.version').Count -eq 1) -and ((Get-Item -LiteralPath ($sharedLedger + '.version')).Length -lt 256)) 'marker shape'
$lock2.Release()
Ok 'releasing removes only the exact lock file' (-not (Test-Path -LiteralPath ($sharedLedger + '.lock'))) 'no lock remains'

# ---------------------------------------------------------------------------
Write-Output ''
Write-Output ("{0} passed, {1} failed" -f $script:pass, $script:fail)
Write-Output ''
if ($script:fail -gt 0) { Write-Output 'RESULT: FAIL'; exit 1 }
Write-Output 'RESULT: PASS'
