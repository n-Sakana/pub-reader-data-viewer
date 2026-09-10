param([string]$Root='', [Parameter(Mandatory=$true)][string]$Evidence)
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.Encoding]::UTF8
if(-not $Root){$Root=Split-Path -Parent $PSScriptRoot}
if(Test-Path -LiteralPath $Evidence){throw 'Use a new evidence directory'}
[IO.Directory]::CreateDirectory($Evidence)|Out-Null
$utf8=New-Object Text.UTF8Encoding($false)
$dataDir=Join-Path $Evidence 'data'
Copy-Item -LiteralPath (Join-Path $Root 'tests/fixtures/sample-v4') -Destination $dataDir -Recurse
$config=Join-Path $Evidence 'settings.json'
$original=Join-Path $Evidence 'original-settings.json'
Copy-Item -LiteralPath (Join-Path $dataDir 'settings.json') -Destination $config
$json=Get-Content -LiteralPath $config -Raw -Encoding UTF8|ConvertFrom-Json
$json.data.ledger.PSObject.Properties.Remove('protectStates')
[IO.File]::WriteAllText($original,($json|ConvertTo-Json -Depth 50),$utf8)
. (Join-Path $Root 'build/test_support.ps1')
Import-RdvProduct -Root $Root
$oldCfg=[Rdv3Config]::Load($original)
$source=[Rdv3Ledger]::BuildFromCsv($oldCfg.Data,$dataDir)
$states=[string[]]@($source.Lines|ForEach-Object {'FALSE'})
for($i=0;$i -lt 21;$i++){$states[$i]='TRUE'}
$oldPath=Join-Path $Evidence 'legacy.xlsx';$output=Join-Path $Evidence 'migrated.xlsx'
[Rdv3Xlsx]::Write($oldPath,$oldCfg.Data.Head,$oldCfg.Screen.Work.Column,$source.Lines,$states,'legacy-fixture',[Rdv3Files]::LegacyStorageContract($oldCfg.Data,$oldCfg.Screen.Work))
$oldHash=(Get-FileHash -LiteralPath $oldPath).Hash
function Invoke-Migration([string]$Name,[string]$Origin,[string]$Destination,[bool]$Confirm,[int]$Expected) {
    $args=@('-NoProfile','-ExecutionPolicy','Bypass','-File',(Join-Path $Root 'tools/Migrate-Ledger.ps1'),'-OriginalConfig',$Origin,'-Config',$config,'-DataDir',$dataDir,'-Ledger',$oldPath,'-Output',$Destination)
    if($Confirm){$args+='-ConfirmOriginalDefinition'}
    $start=New-Object Diagnostics.ProcessStartInfo
    $start.FileName='powershell.exe';$start.Arguments=($args|ForEach-Object {'"'+$_+'"'}) -join ' '
    $start.UseShellExecute=$false;$start.CreateNoWindow=$true;$start.RedirectStandardOutput=$true;$start.RedirectStandardError=$true
    $start.StandardOutputEncoding=[Text.Encoding]::GetEncoding(932);$start.StandardErrorEncoding=[Text.Encoding]::GetEncoding(932)
    $process=[Diagnostics.Process]::Start($start)
    $stdout=$process.StandardOutput.ReadToEndAsync();$stderr=$process.StandardError.ReadToEndAsync()
    $process.WaitForExit();Write-Output $stdout.Result;Write-Output $stderr.Result
    Write-Output ('Closed migration command PID '+$process.Id+' exit='+$process.ExitCode)
    if($process.ExitCode -ne $Expected){throw ($Name+' exit='+$process.ExitCode+' expected='+$Expected)}
    if((Get-FileHash -LiteralPath $oldPath).Hash -ne $oldHash){throw ($Name+' changed source ledger')}
    Write-Output ('PASS '+$Name+' exit='+$Expected+'; original XLSX unchanged')
}
Invoke-Migration 'no operator attestation' $original $output $false 3
if(Test-Path -LiteralPath $output){throw 'Unconfirmed migration created output'}
$changed=Join-Path $Evidence 'different-original.json'
[IO.File]::WriteAllText($changed,[IO.File]::ReadAllText($original).Replace('[0-9]{8}','[0-9]{7}'),$utf8)
Invoke-Migration 'different original definition' $changed $output $true 3
if(Test-Path -LiteralPath $output){throw 'Definition mismatch created output'}
$shared=[Rdv3SharedFiles]::new($oldPath,'migration-test','tester','test')
$owner=$null;$lease=$shared.TryAcquire([ref]$owner)
if($null -eq $lease){throw 'Could not establish lock precondition'}
try {Invoke-Migration 'shared source lock held' $original $output $true 3} finally {$lease.Dispose()}
Invoke-Migration 'migration to distinct new file' $original $output $true 0
$newCfg=[Rdv3Config]::Load($config)
[Rdv3Ledger]::BuildFromCsv($newCfg.Data,$dataDir)|Out-Null
$savedLines=[string[]]@();$savedStates=[string[]]@();$warning='';$protection=$null
[Rdv3Xlsx]::ReadProtected($output,$newCfg.Data.Head,$newCfg.Screen.Work.Column,[ref]$savedLines,[ref]$savedStates,[ref]$warning,[Rdv3Files]::StorageContract($newCfg.Data,$newCfg.Screen.Work),[Rdv3Files]::LegacyStorageContract($newCfg.Data,$newCfg.Screen.Work),[Rdv3BusinessDefinition]::Bound($newCfg.Data),[ref]$protection)
if($protection.Legacy -or $warning -or -not [Rdv3Ledger]::SameLedger($source.Lines,$states,$savedLines,$savedStates)){throw 'Migrated data/state/definition mismatch'}
Write-Output 'PASS migrated real XLSX: 100 full rows, 21 saved done states, fixed definition, empty archive'
$outputHash=(Get-FileHash -LiteralPath $output).Hash
Invoke-Migration 'existing destination' $original $output $true 3
if((Get-FileHash -LiteralPath $output).Hash -ne $outputHash){throw 'Existing destination changed'}
Invoke-Migration 'source used as destination' $original $oldPath $true 3
Write-Output 'PASS migration CLI checks complete'
