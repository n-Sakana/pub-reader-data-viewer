param([Parameter(Mandatory=$true)][string]$BaselineRoot,[string]$Evidence='', [switch]$Snapshot,[string]$ProductRoot='')
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
if(-not $Evidence){$Evidence=Join-Path $repo ('work/spec-parity-'+[Guid]::NewGuid().ToString('N'))}
$Evidence=[IO.Path]::GetFullPath($Evidence)
if($Snapshot){
    . (Join-Path $repo 'build/test_support.ps1')
    Import-RdvProduct -Root $ProductRoot
    $configPath=Join-Path $ProductRoot 'configs/sample/settings.json'
    if(-not(Test-Path -LiteralPath $configPath)){$configPath=Join-Path $ProductRoot 'settings.json'}
    $cfg=[Rdv3Config]::Load($configPath)
    $april=Join-Path $repo 'samples/current/data'
    $may=Join-Path $repo 'samples/current/追加CSVデータ/5月分（3件追加）'
    if(-not(Test-Path -LiteralPath $may)){throw 'Missing published May sample'}
    $initial=[Rdv3Process]::Run($cfg.Data,$cfg.Data.UpdateJob,$april,[string[]]@(),[string[]]@(),$cfg.Screen.Work.InitialStored)
    $fields=[Rdv3Fields]::new($initial.Columns)
    $judgments=@()
    foreach($line in $initial.Lines){
        $view=[Rdv3View]::new();$view.Record=$line.Split([char]9)
        $judgments+=[Rdv3Eval]::Judge($cfg.Screen.Judgments['paymentStatus'],$view,$fields).Result.Id
    }
    $job=$cfg.Data.JobOf('delete-processed-records')
    $unconfirmed=[Rdv3Process]::Run($cfg.Data,$job,$april,$initial.Lines,$initial.States,$cfg.Screen.Work.InitialStored)
    $states=[string[]]$initial.States.Clone();$states[0]='TRUE'
    $deleted=[Rdv3Process]::Run($cfg.Data,$job,$april,$initial.Lines,$states,$cfg.Screen.Work.InitialStored)
    $protection=[Rdv3LedgerProtection]::Create($cfg.Data)
    $protection.ArchiveRemoved($cfg.Data,$initial.Lines,$states,$deleted.Lines)
    $updated=[Rdv3Process]::Run($cfg.Data,$cfg.Data.UpdateJob,$may,$deleted.Lines,$deleted.States,$cfg.Screen.Work.InitialStored)
    $maySource=[Rdv3Process]::Run($cfg.Data,$cfg.Data.UpdateJob,$may,[string[]]@(),[string[]]@(),$cfg.Screen.Work.InitialStored)
    $protection.ProtectUpdate($cfg.Data,$cfg.Screen.Work.InitialStored,$deleted.Lines,$deleted.States,$maySource.Lines,$updated.Update)
    $reimport=[Rdv3Process]::Run($cfg.Data,$cfg.Data.UpdateJob,$april,$deleted.Lines,$deleted.States,$cfg.Screen.Work.InitialStored)
    $protection.ProtectUpdate($cfg.Data,$cfg.Screen.Work.InitialStored,$deleted.Lines,$deleted.States,$initial.Lines,$reimport.Update)
    $kept=[Rdv3Process]::Run($cfg.Data,$cfg.Data.UpdateJob,$may,$initial.Lines,$states,$cfg.Screen.Work.InitialStored)
    [Rdv3LedgerProtection]::Create($cfg.Data).ProtectUpdate($cfg.Data,$cfg.Screen.Work.InitialStored,$initial.Lines,$states,$maySource.Lines,$kept.Update)
    if($initial.Lines.Length -ne 3 -or $unconfirmed.Deleted -ne 0 -or $deleted.Deleted -ne 1 -or $updated.Update.Lines.Length -ne 5 -or $reimport.Update.SkippedDeleted -ne 1){throw 'Required 3/2/5 or saved-state/deleted-record behavior differs'}
    if($kept.Update.Lines -cnotcontains $initial.Lines[0]){throw 'Saved confirmed full row was not retained'}
    $result=[ordered]@{stateColumn=$cfg.Screen.Work.Column;contract=[Rdv3Files]::StorageContract($cfg.Data,$cfg.Screen.Work);definition=$protection.Definition;columns=$initial.Columns;initialLines=$initial.Lines;initialStates=$initial.States;judgments=$judgments;unconfirmedDeleted=$unconfirmed.Deleted;deletedLines=$deleted.Lines;deletedStates=$deleted.States;archiveLine=$protection.Deleted[0].Line;archiveState=$protection.Deleted[0].State;updatedLines=$updated.Update.Lines;updatedStates=$updated.Update.States;reimportLines=$reimport.Update.Lines;skippedDeleted=$reimport.Update.SkippedDeleted;protectedLines=$kept.Update.Lines;protectedStates=$kept.Update.States}
    [IO.File]::WriteAllText($Evidence,($result|ConvertTo-Json -Depth 30),[Text.UTF8Encoding]::new($false))
    return
}
if(Test-Path -LiteralPath $Evidence){throw 'Use a new evidence directory'}
[IO.Directory]::CreateDirectory($Evidence)|Out-Null
$hostExe=Join-Path $PSHOME 'powershell.exe'
foreach($case in @(@{name='baseline';root=$BaselineRoot},@{name='current';root=$repo})){
    & $hostExe -NoProfile -STA -ExecutionPolicy Bypass -File $PSCommandPath -BaselineRoot $BaselineRoot -Snapshot -ProductRoot $case.root -Evidence (Join-Path $Evidence ($case.name+'.json'))
    if($LASTEXITCODE -ne 0){throw ('Snapshot failed: '+$case.name)}
}
$before=Get-Content -LiteralPath (Join-Path $Evidence 'baseline.json') -Raw -Encoding UTF8|ConvertFrom-Json
$after=Get-Content -LiteralPath (Join-Path $Evidence 'current.json') -Raw -Encoding UTF8|ConvertFrom-Json
if($before.stateColumn -cne '処理状態' -or $after.stateColumn -cne '確認状態' -or $before.contract -ceq $after.contract){throw 'Approved heading and storage-contract change did not occur'}
foreach($snapshotValue in @($before,$after)){$snapshotValue.PSObject.Properties.Remove('stateColumn');$snapshotValue.PSObject.Properties.Remove('contract')}
if(($before|ConvertTo-Json -Depth 30) -cne ($after|ConvertTo-Json -Depth 30)){throw 'Baseline/current business snapshots differ'}
'PASS approved heading/contract change only; definition, all rows/states, payment judgments, 3/2/5, full-row protection and deleted-record reimport exclusion are identical'
