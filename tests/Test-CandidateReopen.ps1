param([string]$Root='', [Parameter(Mandatory=$true)][string]$Evidence, [string]$NodePath='')
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.Encoding]::UTF8
if(-not $Root){$Root=Split-Path -Parent $PSScriptRoot}
if(-not $NodePath){$NodePath=(Get-Command node -ErrorAction Stop).Source}
$utf8=New-Object Text.UTF8Encoding($false)
Add-Type -AssemblyName UIAutomationClient,UIAutomationTypes,System.Windows.Forms
Add-Type -Path (Join-Path $PSScriptRoot 'NativeWindowProbe.cs')
[ReaderWindowProbe]::PhysicalPixels()
$scratch=Join-Path $Evidence 'app'
if(Test-Path -LiteralPath $Evidence){throw 'Use a new evidence directory'}
[IO.Directory]::CreateDirectory($scratch)|Out-Null
foreach($name in @('src','web','lib')){Copy-Item -LiteralPath (Join-Path $Root $name) -Destination $scratch -Recurse}
Copy-Item -LiteralPath (Join-Path $Root 'tests/fixtures/sample-v4') -Destination (Join-Path $scratch 'data') -Recurse
$settings=Join-Path $scratch 'settings.json'
$json=Get-Content -LiteralPath (Join-Path $Root 'tests/fixtures/sample-v4/settings.json') -Raw -Encoding UTF8|ConvertFrom-Json
$json.watch.targets=@()
[IO.File]::WriteAllText($settings,($json|ConvertTo-Json -Depth 50),$utf8)
. (Join-Path $Root 'build/test_support.ps1')
Import-RdvProduct -Root $Root
$cfg=[Rdv3Config]::Load($settings)
[Rdv3Ledger]::BuildFromCsv($cfg.Data,(Join-Path $scratch 'data'))|Out-Null
$result=[Rdv3Process]::Run($cfg.Data,$cfg.Data.UpdateJob,(Join-Path $scratch 'data'),[string[]]@(),[string[]]@(),$cfg.Screen.Work.InitialStored)
$fields=[Rdv3Fields]::new($result.Columns)
$card=$fields.IndexOf('PAY.会員番号照合用')
$lines=[string[]]$result.Lines.Clone()
foreach($i in @(0,2)){$cells=$lines[$i+1].Split([char]9);$cells[$card]=$lines[$i].Split([char]9)[$card];$lines[$i+1]=[string]::Join("`t",$cells)}
if ('Rdv3LedgerProtection' -as [type]) {
    [Rdv3Xlsx]::Write((Join-Path $scratch $cfg.Ledger),$cfg.Data.Head,$cfg.Screen.Work.Column,$lines,$result.States,'candidate-reopen-test',[Rdv3Files]::StorageContract($cfg.Data,$cfg.Screen.Work),[Rdv3LedgerProtection]::Create($cfg.Data))
} else {
    [Rdv3Xlsx]::Write((Join-Path $scratch $cfg.Ledger),$cfg.Data.Head,$cfg.Screen.Work.Column,$lines,$result.States,'candidate-reopen-test',[Rdv3Files]::StorageContract($cfg.Data,$cfg.Screen.Work))
}
$listener=New-Object Net.Sockets.TcpListener([Net.IPAddress]::Loopback,0)
$listener.Start();$port=$listener.LocalEndpoint.Port;$listener.Stop()
[IO.File]::WriteAllText((Join-Path $Evidence 'case.json'),(@{port=$port;paidKey=$lines[0].Split([char]9)[$card];otherKey=$lines[2].Split([char]9)[$card]}|ConvertTo-Json),$utf8)
$start=New-Object Diagnostics.ProcessStartInfo
$start.FileName='powershell.exe';$start.Arguments='-NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "'+(Join-Path $scratch 'src/ReaderDataViewer.ps1')+'"'
$start.WorkingDirectory=$scratch;$start.UseShellExecute=$false;$start.CreateNoWindow=$true
$start.EnvironmentVariables['WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS']='--remote-debugging-port='+$port
$reader=[Diagnostics.Process]::Start($start)
$driver=$null;$owned=@($reader.Id)
try {
    $start=New-Object Diagnostics.ProcessStartInfo
    $start.FileName=$NodePath;$start.Arguments='"'+(Join-Path $PSScriptRoot 'candidate-reopen.cjs')+'" "'+$Evidence+'" "'+$Root+'"'
    $start.UseShellExecute=$false;$start.CreateNoWindow=$true;$start.RedirectStandardOutput=$true;$start.RedirectStandardError=$true
    $driver=[Diagnostics.Process]::Start($start)
    $stdout=$driver.StandardOutput.ReadToEndAsync();$stderr=$driver.StandardError.ReadToEndAsync()
    $handled=0;$deadline=[DateTime]::UtcNow.AddSeconds(150)
    while(-not $driver.HasExited -and [DateTime]::UtcNow -lt $deadline){
        $all=@(Get-CimInstance Win32_Process -Filter "Name='powershell.exe' OR Name='msedgewebview2.exe'")
        do{$count=$owned.Count;foreach($p in $all){if($owned -contains [int]$p.ParentProcessId -and $owned -notcontains [int]$p.ProcessId){$owned+=[int]$p.ProcessId}}}while($count -ne $owned.Count)
        $requestPath=Join-Path $Evidence 'native-request.json'
        if(Test-Path -LiteralPath $requestPath){
            try{$request=Get-Content -LiteralPath $requestPath -Raw|ConvertFrom-Json}catch{$request=$null}
            if($request -and $request.id -gt $handled){
                $until=[DateTime]::UtcNow.AddSeconds(7)
                do {
                    $windows=@([ReaderWindowProbe]::Read([int[]]@($reader.Id))|Where-Object {$_.visible -and $_.title})
                    $main=@($windows|Where-Object {$_.title -eq 'Reader Data Viewer'})
                    $dialogs=@($windows|Where-Object {$_.title -ne 'Reader Data Viewer'})
                    $enabled=$false
                    if($main.Count -eq 1){$enabled=[System.Windows.Automation.AutomationElement]::FromHandle([IntPtr]$main[0].handle).Current.IsEnabled}
                    $onScreen=$false
                    if($dialogs.Count -eq 1){
                        $d=$dialogs[0]
                        foreach($screen in [Windows.Forms.Screen]::AllScreens){$b=$screen.Bounds;if($d.left -ge $b.Left -and $d.top -ge $b.Top -and $d.right -le $b.Right -and $d.bottom -le $b.Bottom){$onScreen=$true}}
                    }
                    $pass=if($request.kind -eq 'dialog'){(-not $enabled) -and $onScreen}else{$enabled -and $dialogs.Count -eq 0}
                    if(-not $pass){Start-Sleep -Milliseconds 100}
                }while(-not $pass -and [DateTime]::UtcNow -lt $until)
                $reply=@{id=$request.id;kind=$request.kind;cycle=$request.cycle;pid=$reader.Id;utc=[DateTime]::UtcNow.ToString('o');pass=[bool]$pass;parentEnabled=[bool]$enabled;dialogOnScreen=[bool]$onScreen;windows=$windows}
                [IO.File]::WriteAllText((Join-Path $Evidence 'native-reply.json'),($reply|ConvertTo-Json -Depth 5),$utf8)
                $handled=$request.id
            }
        }
        Start-Sleep -Milliseconds 60
    }
    if(-not $driver.HasExited){throw 'Candidate test timed out'}
    $driver.WaitForExit()
    [IO.File]::WriteAllText((Join-Path $Evidence 'stdout.txt'),$stdout.Result,$utf8)
    [IO.File]::WriteAllText((Join-Path $Evidence 'stderr.txt'),$stderr.Result,$utf8)
    Write-Output $stdout.Result
    Write-Output $stderr.Result
    if($driver.ExitCode -ne 0){throw ('Candidate test failed: '+$driver.ExitCode)}
    if(-not $reader.WaitForExit(10000)){throw 'Reader did not exit'}
    Write-Output ('PASS native candidate reopen: Reader PID '+$reader.Id+' exited '+$reader.ExitCode)
}finally{
    if($driver -and -not $driver.HasExited){Stop-Process -Id $driver.Id -Force}
    foreach($ownedId in $owned){if(Get-Process -Id $ownedId -ErrorAction SilentlyContinue){Stop-Process -Id $ownedId -Force -ErrorAction SilentlyContinue}}
    [IO.File]::WriteAllText((Join-Path $Evidence 'owned-pids.json'),($owned|ConvertTo-Json),$utf8)
}
