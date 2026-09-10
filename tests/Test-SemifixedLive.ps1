param([string]$Root='', [string]$Evidence='', [string]$NodePath='')
$ErrorActionPreference='Stop'
if(-not $NodePath){$NodePath=(Get-Command node -ErrorAction Stop).Source}
[Console]::OutputEncoding=[Text.Encoding]::UTF8
if(-not $Root){$Root=Split-Path -Parent $PSScriptRoot}
if(-not $Evidence){$Evidence=Join-Path $Root ('work/semifixed/live-'+[Guid]::NewGuid().ToString('N'))}
$scratch=Join-Path $Evidence 'app'
[IO.Directory]::CreateDirectory($scratch)|Out-Null
Copy-Item -LiteralPath (Join-Path $Root 'settings.json') -Destination $scratch
foreach($name in @('src','web','lib')){Copy-Item -LiteralPath (Join-Path $Root $name) -Destination $scratch -Recurse}
Copy-Item -LiteralPath (Join-Path $Root 'tests/fixtures/sample-v4') -Destination (Join-Path $scratch 'data') -Recurse
[IO.Directory]::CreateDirectory((Join-Path $scratch 'output'))|Out-Null
$settings=Join-Path $scratch 'settings.json'
$json=[IO.File]::ReadAllText($settings)|ConvertFrom-Json
$json.watch.targets=@()
$utf8=New-Object Text.UTF8Encoding($false)
[IO.File]::WriteAllText($settings,($json|ConvertTo-Json -Depth 50),$utf8)
. (Join-Path $Root 'build/test_support.ps1')
Import-RdvProduct -Root $Root
$cfg=[Rdv3Config]::Load($settings)
$result=[Rdv3Process]::Run($cfg.Data,$cfg.Data.UpdateJob,(Join-Path $scratch 'data'),[string[]]@(),[string[]]@(),$cfg.Screen.Work.InitialStored)
$fields=[Rdv3Fields]::new($result.Columns)
$card=$fields.IndexOf('PAY.会員番号照合用')
$first=$result.Lines[0].Split([char]9)
$lines=[string[]]$result.Lines.Clone()
$second=$lines[1].Split([char]9);$second[$card]=$first[$card];$lines[1]=[string]::Join("`t",$second)
$ledger=Join-Path $scratch $cfg.Ledger
[Rdv3Xlsx]::Write($ledger,$cfg.Data.Head,$cfg.Screen.Work.Column,$lines,$result.States,'fixed-screen-test',[Rdv3Files]::StorageContract($cfg.Data,$cfg.Screen.Work))
$pay=$fields.IndexOf('PAYMAP.決済確認済')
$missingKey=@($result.Lines | Where-Object { $_.Split([char]9)[$pay] -eq '' })[0].Split([char]9)[$card]
$appColumn=$fields.IndexOf('APP.申請番号')
$noAppKey=@($result.Lines | Where-Object { $_.Split([char]9)[$appColumn] -eq '' })[0].Split([char]9)[$card]
$case=@{paidKey=$first[$card];missingKey=$missingKey;noAppKey=$noAppKey;ledger=$ledger;settings=$settings;scratch=$scratch}
[IO.File]::WriteAllText((Join-Path $Evidence 'case.json'),($case|ConvertTo-Json),$utf8)
$listener=New-Object Net.Sockets.TcpListener([Net.IPAddress]::Loopback,0);$listener.Start();$port=$listener.LocalEndpoint.Port;$listener.Stop()
$start=New-Object Diagnostics.ProcessStartInfo
$start.FileName='powershell.exe';$start.Arguments='-NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "'+(Join-Path $scratch 'src/ReaderDataViewer.ps1')+'"'
$start.WorkingDirectory=$scratch;$start.UseShellExecute=$false;$start.CreateNoWindow=$true
$start.EnvironmentVariables['WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS']='--remote-debugging-port='+$port
$process=[Diagnostics.Process]::Start($start)
Write-Output ('Owned Reader PID '+$process.Id+'; evidence='+$Evidence)
try{
    & $NodePath (Join-Path $Root 'tests/semifixed-live.cjs') $port $Evidence
    if($LASTEXITCODE -ne 0){throw ('UI verification failed: '+$LASTEXITCODE)}
    if(-not $process.WaitForExit(10000)){throw 'Reader did not close'}
    Write-Output ('Closed PID '+$process.Id+'; exit='+$process.ExitCode)
}finally{if(-not $process.HasExited){Stop-Process -Id $process.Id -Force}}
$savedLines=[string[]]@();$savedStates=[string[]]@()
[Rdv3Xlsx]::Read((Join-Path $Evidence 'sent-ledger.xlsx'),$cfg.Data.Head,$cfg.Screen.Work.Column,[ref]$savedLines,[ref]$savedStates)
if($savedLines.Length -ne 100 -or @($savedStates|Where-Object {$_ -eq 'TRUE'}).Count -ne 1){throw 'The sent state did not persist to the actual ledger'}
[Rdv3Xlsx]::Read((Join-Path $Evidence 'deleted-ledger.xlsx'),$cfg.Data.Head,$cfg.Screen.Work.Column,[ref]$savedLines,[ref]$savedStates)
if($savedLines.Length -ne 99){throw 'Processed-only deletion did not persist to the actual ledger'}
Write-Output 'PASS actual ledger readback: sent state=1, deleted rows=1, unprocessed matching rows kept, remaining=99'
