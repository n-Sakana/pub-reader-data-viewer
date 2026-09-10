param([string]$Root='')
$ErrorActionPreference='Stop'
if(-not $Root){$Root=Split-Path -Parent $PSScriptRoot}
. (Join-Path $Root 'build/test_support.ps1')
Import-RdvProduct -Root $Root
$scratch=Join-Path $Root ('work/semifixed/config-'+[Guid]::NewGuid().ToString('N'))
[IO.Directory]::CreateDirectory($scratch)|Out-Null
$utf8=New-Object Text.UTF8Encoding($false)
$legacy=[Rdv3Config]::Load((Join-Path $Root 'tests/fixtures/legacy-payment-settings.json'))
$after=[Rdv3Config]::Load((Join-Path $Root 'configs/sample/settings.json'))
if([Rdv3Files]::LegacyStorageContract($legacy.Data,$legacy.Screen.Work) -ne [Rdv3Files]::LegacyStorageContract($after.Data,$after.Screen.Work)){throw 'The two versions disagree on row/state storage'}
if([Rdv3Files]::StorageContract($legacy.Data,$legacy.Screen.Work) -eq [Rdv3Files]::StorageContract($after.Data,$after.Screen.Work)){throw 'Protection format was not separated from legacy'}
$matching=Get-Content -LiteralPath (Join-Path $Root 'tests/fixtures/legacy-payment-settings.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$matching.data.ledger | Add-Member NoteProperty protectStates @('TRUE')
$matchingPath=Join-Path $scratch 'matching-main-settings.json'
[IO.File]::WriteAllText($matchingPath,($matching|ConvertTo-Json -Depth 50),$utf8)
$main=[Rdv3Config]::Load($matchingPath)
if([Rdv3Files]::StorageContract($main.Data,$main.Screen.Work) -ne [Rdv3Files]::StorageContract($after.Data,$after.Screen.Work)){throw 'Protected main and fixed variants disagree on storage'}
$old=[Rdv3Config]::Load((Join-Path $Root 'tests/fixtures/legacy-app-settings.json'))
if([Rdv3Files]::StorageContract($old.Data,$old.Screen.Work) -eq [Rdv3Files]::StorageContract($after.Data,$after.Screen.Work)){throw 'Old APP contract was reused'}
if($old.Ledger -eq $after.Ledger){throw 'Old APP ledger path was reused'}
Write-Output 'PASS matching variant storage; new ledger isolated from old APP storage'
$source=[IO.File]::ReadAllText((Join-Path $Root 'settings.json'))
foreach($case in @('missing binding','unknown field','extra layout','wrong job kind')){
    $config=$source|ConvertFrom-Json
    switch($case){
        'missing binding' {$config.screen.bindings.PSObject.Properties.Remove('remarks')}
        'unknown field' {$config.screen.bindings.remarks.field='APP.DOES_NOT_EXIST'}
        'extra layout' {$config.screen|Add-Member NoteProperty card @{width=600}}
        'wrong job kind' {$config.screen.actions.updateRecords=$config.screen.actions.deleteRecords}
    }
    $path=Join-Path $scratch ($case+'.json')
    [IO.File]::WriteAllText($path,($config|ConvertTo-Json -Depth 50),$utf8)
    $rejected=$false
    try{[Rdv3Config]::Load($path)|Out-Null}catch{$rejected=$true}
    if(-not $rejected){throw ('Invalid configuration accepted: '+$case)}
    Write-Output ('PASS rejected: '+$case)
}
