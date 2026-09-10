param([string]$Root='', [string]$Config='', [string]$DataDir='', [string]$Evidence='', [string]$ActualConfig='')
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.Encoding]::UTF8
if(-not $Root){$Root=Split-Path -Parent $PSScriptRoot}
if(-not $DataDir){$DataDir=Join-Path $Root 'tests/fixtures/sample-v4'}
if(-not $Config){$Config=Join-Path $DataDir 'settings.json'}
if(-not $ActualConfig){$ActualConfig=Join-Path $Root 'configs/sample/settings.json'}
. (Join-Path $Root 'build/test_support.ps1')
Import-RdvProduct -Root $Root
$cfg=[Rdv3Config]::Load($Config)
$data=$DataDir
if(-not $Evidence){$Evidence=Join-Path $Root ('work/payment-basis/check-'+[Guid]::NewGuid().ToString('N'))}
[IO.Directory]::CreateDirectory($evidence)|Out-Null
function Assert($Condition,[string]$Message){if(-not $Condition){throw $Message}}
$headless=[Rdv3Headless]::Evaluate($cfg,$Root,$data,$true,'')|ConvertFrom-Json
Assert ($headless.summary.rows -eq 100 -and $headless.warnings.Count -eq 0) 'Headless count or input warning differs'
foreach($input in $headless.inputs){Assert ($input.rows -eq 100) ('Input count: '+$input.id)}
function Run-Update([string]$Directory){[Rdv3Process]::Run($cfg.Data,$cfg.Data.UpdateJob,$Directory,[string[]]@(),[string[]]@(),$cfg.Screen.Work.InitialStored)}
$result=Run-Update $data
Assert ($result.ValueOf('PAYMAP').Count -eq 100) 'MAP/PAY pair count differs'
Assert ($result.Joins[0].UnmatchedLeft -eq 0 -and $result.Joins[0].UnmatchedRight -eq 0) 'Initial MAP/PAY should all pair'
Assert ($result.Joins[1].UnmatchedLeft -eq 20 -and $result.Joins[1].UnmatchedRight -eq 20) 'APP overlap differs'
$fields=[Rdv3Fields]::new($result.Columns)
$app=$fields.IndexOf('APP.申請番号')
$pay=$fields.IndexOf('PAY.オーダーID')
$map=$fields.IndexOf('MAP.【共通】取引ID')
$card=$fields.IndexOf('PAY.会員番号照合用')
$number=$fields.IndexOf('PAY.申請番号照合用')
$flag=$fields.IndexOf('PAYMAP.決済確認済')
$counts=@{paid=0;unpaid=0}
$noApp=0
$identities=[Collections.Generic.HashSet[string]]::new()
$masked=@($cfg.Screen.AllBindings()|Where-Object {$_.Requires -contains 'PAYMAP.決済確認済'})
Assert ($masked.Count -eq 9) 'Nine application detail fields must remain conditional'
foreach($line in $result.Lines){
    $row=$line.Split([char]9)
    Assert ($row[$pay] -ne '' -and $row[$map] -ne '') 'A one-sided row entered the ledger'
    Assert ($identities.Add($row[$pay])) 'Duplicate transaction identity'
    Assert ($row[$card] -ne '' -and $row[$number] -ne '') 'A row without APP cannot be searched'
    $view=[Rdv3View]::new();$view.Record=$row
    $verdict=[Rdv3Eval]::Judge($cfg.Screen.Judgments['paymentStatus'],$view,$fields).Result.Id
    Assert ($counts.ContainsKey($verdict)) ('Unexpected verdict '+$verdict)
    $counts[$verdict]++
    foreach($binding in $masked){
        $text=[Rdv3Eval]::Evaluate($binding,$view,$fields,$cfg.Screen.Work).Text
        if($verdict -eq 'unpaid'){Assert ($text -eq '') 'Unpaid application detail was displayed'}
        elseif($row[$app] -ne ''){Assert ($text -ne '') 'Paid application detail was hidden'}
    }
    if($row[$app] -eq ''){$noApp++;Assert ($verdict -eq 'paid') 'APP absence incorrectly made a paid pair unpaid'}
}
Assert ($result.Lines.Length -eq 100 -and $noApp -eq 20) 'PAYMAP basis must keep 20 rows lacking APP'
Write-Output 'PASS paired input: 100 rows, including 20 without APP; all have searchable keys'
Assert ($counts.paid -eq 80 -and $counts.unpaid -eq 20) 'Payment AND must yield 80 paid and 20 unpaid'
Write-Output 'PASS payment AND: 80 paid, 10 MAP-only not paid, 10 PAY-only not paid'
$repeat=[Rdv3Process]::Run($cfg.Data,$cfg.Data.UpdateJob,$data,$result.Lines,$result.States,$cfg.Screen.Work.InitialStored)
Assert ($repeat.Lines.Length -eq 100) 'Repeated input duplicated rows'
Write-Output 'PASS repeated import: 100 rows'
$delete=@($cfg.Data.Jobs|Where-Object {$_.Kind -eq 'delete'})[0]
$unprocessed=[Rdv3Process]::Run($cfg.Data,$delete,$data,$result.Lines,$result.States,$cfg.Screen.Work.InitialStored)
Assert ($unprocessed.Deleted -eq 0 -and $unprocessed.Lines.Length -eq 100) 'Unprocessed rows were deleted'
Assert ($unprocessed.ValueOf('DEL').Count -eq 100) 'Deletion input count differs'
$states=[string[]]$result.States.Clone()
for($i=0;$i -lt 10;$i++){$states[$i]='TRUE'}
$states[30]='TRUE' # processed but absent from the deletion list
$mixed=[Rdv3Process]::Run($cfg.Data,$delete,$data,$result.Lines,$states,$cfg.Screen.Work.InitialStored)
Assert ($mixed.Deleted -eq 10 -and $mixed.Lines.Length -eq 90) 'Pair match AND processed state failed'
Assert (@($mixed.States|Where-Object {$_ -eq 'TRUE'}).Count -eq 1) 'Unmatched processed state was lost'
for($i=10;$i -lt 20;$i++){$states[$i]='TRUE'}
$deleted=[Rdv3Process]::Run($cfg.Data,$delete,$data,$result.Lines,$states,$cfg.Screen.Work.InitialStored)
Assert ($deleted.Deleted -eq 20 -and $deleted.Lines.Length -eq 80) 'Expected 20 deleted, 80 remaining'
Write-Output 'PASS deletion: unprocessed=0, half processed=10, all matching processed=20; unmatched processed kept'
$repeatedDelete=[Rdv3Process]::Run($cfg.Data,$delete,$data,$deleted.Lines,$deleted.States,$cfg.Screen.Work.InitialStored)
Assert ($repeatedDelete.Deleted -eq 0) 'Repeated deletion removed an unrelated row'
$cross=$result.Lines[0].Split([char]9);$cross[$number]=$result.Lines[1].Split([char]9)[$number]
$crossed=[Rdv3Process]::Run($cfg.Data,$delete,$data,[string[]]@([string]::Join("`t",$cross)),[string[]]@('TRUE'),$cfg.Screen.Work.InitialStored)
Assert ($crossed.Deleted -eq 0) 'Crossed keys were treated as a matched pair'
foreach($stateValue in @('','UNKNOWN','FALSE')){
    $kept=[Rdv3Process]::Run($cfg.Data,$delete,$data,[string[]]@($result.Lines[0]),[string[]]@($stateValue),$cfg.Screen.Work.InitialStored)
    Assert ($kept.Deleted -eq 0) ('Invalid state deleted a row: '+$stateValue)
}
foreach($name in @('MAP','PAY')){
    $copy=Join-Path $evidence ('without-'+$name)
    [IO.Directory]::CreateDirectory($copy)|Out-Null
    foreach($table in $cfg.Data.Tables){Copy-Item -LiteralPath (Join-Path $data $table.File) -Destination $copy}
    $table=$cfg.Data.TableOf($name);$path=Join-Path $copy $table.File
    $lines=[IO.File]::ReadAllLines($path,$table.Enc)
    [IO.File]::WriteAllLines($path,[string[]](@($lines[0])+@($lines[2..($lines.Length-1)])),$table.Enc)
    $boundary=Run-Update $copy
    Assert ($boundary.Lines.Length -eq 99) ('Missing '+$name+' was still included')
    Assert (@($boundary.Lines|Where-Object { $_.Split([char]9)[$pay] -eq 'ORD-AA10000000AA-0000' }).Count -eq 0) 'The unmatched transaction remained'
    Write-Output ('PASS missing '+$name+': 99 rows; unmatched transaction absent')
}
$utf8=New-Object Text.UTF8Encoding($false)
$sourceJson=[IO.File]::ReadAllText($Config)|ConvertFrom-Json
$badDelete=@($sourceJson.data.jobs|Where-Object {$_.kind -eq 'delete'})[0]
$badDelete.inputs+=@{table='MAP'}
$badDelete.steps[1].target1='MAP'
$invalid=Join-Path $evidence 'input-state-invalid.json'
[IO.File]::WriteAllText($invalid,($sourceJson|ConvertTo-Json -Depth 50),$utf8)
$rejected=$false
try{[Rdv3Config]::Load($invalid)|Out-Null}catch{$rejected=$true}
Assert $rejected 'Application state was allowed on an input table'
$actual=[Rdv3Config]::Load($ActualConfig)
$actual.Screen.Check($actual.Data)
Assert ($actual.Data.IdentityRefs[0] -eq 'PAY.オーダーID') 'Supplied config identity is not PAY based'
Assert ($actual.Screen.Judgments['paymentStatus'].Source.Fields[0] -eq 'PAYMAP.決済確認済') 'Supplied config verdict does not use both statuses'
$actualDelete=@($actual.Data.Jobs|Where-Object {$_.Kind -eq 'delete'})[0]
Assert ($actualDelete.Steps[1].Where.Column -eq '$work' -and $actualDelete.Steps[1].Where.Value -eq 'TRUE') 'Supplied config lacks processed-only deletion'
[Rdv3Xlsx]::Write((Join-Path $evidence 'ledger-after.xlsx'),$cfg.Data.Head,$cfg.Screen.Work.Column,$deleted.Lines,$deleted.States,'payment-check',[Rdv3Files]::StorageContract($cfg.Data,$cfg.Screen.Work))
$reloadLines=[string[]]@();$reloadStates=[string[]]@()
[Rdv3Xlsx]::Read((Join-Path $evidence 'ledger-after.xlsx'),$cfg.Data.Head,$cfg.Screen.Work.Column,[ref]$reloadLines,[ref]$reloadStates)
Assert ($reloadLines.Length -eq 80 -and @($reloadStates|Where-Object {$_ -eq 'TRUE'}).Count -eq 1) 'Saved ledger lost retained state'
$report=@{rows=$result.Lines.Length;paid=$counts.paid;unpaid=$counts.unpaid;withoutApplication=$noApp;repeatRows=$repeat.Lines.Length;deleted=$deleted.Deleted;remaining=$deleted.Lines.Length;unprocessedDeleted=$unprocessed.Deleted;mixedDeleted=$mixed.Deleted;crossedPairDeleted=$crossed.Deleted;suppliedConfigChecked=$true}
[IO.File]::WriteAllText((Join-Path $evidence 'verified.json'),($report|ConvertTo-Json),$utf8)
Write-Output ('Evidence: '+$evidence)
