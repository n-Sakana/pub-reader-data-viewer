param([string]$Root='', [Parameter(Mandatory=$true)][string]$Evidence, [string]$NodePath='', [switch]$GuardsOnly, [switch]$AutomaticOnly)
$ErrorActionPreference='Stop'
if($GuardsOnly -and $AutomaticOnly){throw 'Choose either GuardsOnly or AutomaticOnly'}
[Console]::OutputEncoding=[Text.Encoding]::UTF8
if(-not $Root){$Root=Split-Path -Parent $PSScriptRoot}
if(-not $NodePath){$NodePath=(Get-Command node -ErrorAction Stop).Source}
if(Test-Path -LiteralPath $Evidence){throw 'Use a new evidence directory'}
$utf8=New-Object Text.UTF8Encoding($false)
$scratch=Join-Path $Evidence 'app'
[IO.Directory]::CreateDirectory($scratch)|Out-Null
foreach($name in @('src','web','lib')){Copy-Item -LiteralPath (Join-Path $Root $name) -Destination $scratch -Recurse}
Copy-Item -LiteralPath (Join-Path $Root 'tests/fixtures/sample-v4') -Destination (Join-Path $scratch 'data') -Recurse
Add-Type -AssemblyName UIAutomationClient,UIAutomationTypes,System.Windows.Forms,System.Drawing
Add-Type -Path @((Join-Path $PSScriptRoot 'NativeWindowProbe.cs'),(Join-Path $PSScriptRoot 'ProtectionNativeProbe.cs')) -ReferencedAssemblies @('System.Drawing','System.Windows.Forms')
[ReaderWindowProbe]::PhysicalPixels()
$sourceTitle='Reader verification '+[Guid]::NewGuid().ToString('N')
$settings=Join-Path $scratch 'settings.json'
$json=Get-Content -LiteralPath (Join-Path $Root 'tests/fixtures/sample-v4/settings.json') -Raw -Encoding UTF8|ConvertFrom-Json
$json.watch.targets=@(@{enabled=$true;name='Verification source';window=@{name=$sourceTitle;scope='children'};path=@();field=@{automationId='ReaderNumber';controlTypes=@('Edit');requireValuePattern=$true;scope='descendants'};read='value'})
[IO.File]::WriteAllText($settings,($json|ConvertTo-Json -Depth 50),$utf8)
. (Join-Path $Root 'build/test_support.ps1')
Import-RdvProduct -Root $Root
$cfg=[Rdv3Config]::Load($settings)
$merged=[Rdv3Ledger]::BuildFromCsv($cfg.Data,(Join-Path $scratch 'data'))
$fields=[Rdv3Fields]::new($cfg.Data.ColumnRefs)
$card=$fields.IndexOf('PAY.会員番号照合用')
$lines=[string[]]$merged.Lines.Clone()
$states=[string[]]@($lines|ForEach-Object {'FALSE'})
for($i=0;$i -lt 26;$i++){$states[$i]='TRUE'}
$single=$lines[20].Split([char]9)[$card]
$multiple=$lines[24].Split([char]9)[$card]
$cells=$lines[25].Split([char]9);$cells[$card]=$multiple;$lines[25]=[string]::Join("`t",$cells)
$protection=[Rdv3LedgerProtection]::Create($cfg.Data)
$active=[string[]]$lines[20..99];$activeStates=[string[]]$states[20..99]
$protection.ArchiveRemoved($cfg.Data,$lines,$states,$active)
if($AutomaticOnly){
    $lines=[string[]]$merged.Lines.Clone()
    $states=[string[]]@($lines|ForEach-Object {'FALSE'})
    $paid=@();$unpaid=@()
    for($i=0;$i -lt $lines.Length;$i++){
        $view=[Rdv3View]::new();$view.Record=$lines[$i].Split([char]9)
        if([Rdv3Eval]::Judge($cfg.Screen.JudgmentOf('paymentStatus'),$view,$fields).Result.Id -eq 'paid'){$paid+=$i}else{$unpaid+=$i}
    }
    if($paid.Count -ne 80 -or $unpaid.Count -ne 20){throw 'Expected original 80 paid / 20 unpaid sample'}
    $single=$lines[$paid[1]].Split([char]9)[$card]
    $multiple=$lines[$paid[0]].Split([char]9)[$card]
    $cells=$lines[$unpaid[0]].Split([char]9);$cells[$card]=$multiple;$lines[$unpaid[0]]=[string]::Join("`t",$cells)
    $active=$lines;$activeStates=$states
    $protection=[Rdv3LedgerProtection]::Create($cfg.Data)
}
$ledger=Join-Path $scratch $cfg.Ledger
[Rdv3Xlsx]::Write($ledger,$cfg.Data.Head,$cfg.Screen.Work.Column,$active,$activeStates,'native-fixture',[Rdv3Files]::StorageContract($cfg.Data,$cfg.Screen.Work),$protection)
$listener=New-Object Net.Sockets.TcpListener([Net.IPAddress]::Loopback,0);$listener.Start();$port=$listener.LocalEndpoint.Port;$listener.Stop()
$info=@{port=$port;single=$single;multiple=$multiple;manualPending=$lines[26].Split([char]9)[$card];zero='ZZ99999999ZZ';ledger=$ledger;settings=$settings;scratch=$scratch;archivedLine=$lines[0];archivedState=$states[0];fields=$cfg.Data.ColumnRefs;guardsOnly=[bool]$GuardsOnly}
if($AutomaticOnly){
    $info.automaticOnly=$true
    $info.unpaid=$lines[$unpaid[1]].Split([char]9)[$card]
    $info.manualPaid=$lines[$paid[2]].Split([char]9)[$card]
}
[IO.File]::WriteAllText((Join-Path $Evidence 'case.json'),($info|ConvertTo-Json -Depth 5),$utf8)
$sourceExe=Join-Path $Evidence 'ReaderWatchSource.exe'
Add-Type -Path (Join-Path $PSScriptRoot 'ReaderWatchSource.cs') -ReferencedAssemblies @([System.Windows.Window].Assembly.Location,[System.Windows.UIElement].Assembly.Location,[System.Windows.DependencyObject].Assembly.Location,[System.Xaml.XamlReader].Assembly.Location) -OutputAssembly $sourceExe -OutputType WindowsApplication
function Start-Owned([string]$File,[string]$Arguments,[bool]$Redirect=$false) {
    $start=New-Object Diagnostics.ProcessStartInfo
    $start.FileName=$File;$start.Arguments=$Arguments;$start.WorkingDirectory=$scratch
    $start.UseShellExecute=$false;$start.CreateNoWindow=$true
    $start.RedirectStandardOutput=$Redirect;$start.RedirectStandardError=$Redirect
    $start.EnvironmentVariables['WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS']='--remote-debugging-port='+$port
    return [Diagnostics.Process]::Start($start)
}
$source=$null;$reader=$null;$driver=$null;$owned=@()
try {
    $reader=Start-Owned 'powershell.exe' ('-NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "'+(Join-Path $scratch 'src/ReaderDataViewer.ps1')+'"');$owned+=$reader.Id
    $source=Start-Owned $sourceExe ('"'+$sourceTitle+'" '+$reader.Id);$owned+=$source.Id
    Write-Output ('Owned Reader PID '+$reader.Id+'; source PID '+$source.Id)
    $driver=Start-Owned $NodePath ('"'+(Join-Path $PSScriptRoot 'protection-native.cjs')+'" "'+$Evidence+'" "'+$Root+'"') $true
    $owned+=$driver.Id
    $stdout=$driver.StandardOutput.ReadToEndAsync();$stderr=$driver.StandardError.ReadToEndAsync()
    $handled=0;$deadline=[DateTime]::UtcNow.AddMinutes(5)
    while(-not $driver.HasExited -and [DateTime]::UtcNow -lt $deadline){
        $requestPath=Join-Path $Evidence 'native-request.json'
        if(Test-Path -LiteralPath $requestPath){
            try{$request=Get-Content -LiteralPath $requestPath -Raw -Encoding UTF8|ConvertFrom-Json}catch{$request=$null}
            if($request -and $request.id -gt $handled){
                $windows=@([ReaderWindowProbe]::Read([int[]]@($reader.Id))|Where-Object {$_.visible -and $_.title})
                $main=@($windows|Where-Object {$_.title -eq 'Reader Data Viewer'})
                $windowDeadline=[DateTime]::UtcNow.AddSeconds(6)
                while($main.Count -ne 1 -and [DateTime]::UtcNow -lt $windowDeadline){
                    Start-Sleep -Milliseconds 80
                    $windows=@([ReaderWindowProbe]::Read([int[]]@($reader.Id))|Where-Object {$_.visible -and $_.title})
                    $main=@($windows|Where-Object {$_.title -eq 'Reader Data Viewer'})
                }
                if($main.Count -ne 1){throw ('Expected one owned Reader window: '+($windows|ConvertTo-Json -Compress))}
                $mainHandle=[IntPtr]$main[0].handle
                $source.Refresh();$sourceHandle=$source.MainWindowHandle
                if($request.kind -eq 'source'){
                    if($request.minimize){[ProtectionNativeProbe]::ShowWindow($mainHandle,6)|Out-Null}
                    if($request.maximize){[ProtectionNativeProbe]::ShowWindow($mainHandle,3)|Out-Null}
                    [ProtectionNativeProbe]::FocusSource($sourceHandle)
                    $until=[DateTime]::UtcNow.AddSeconds(3)
                    while([ProtectionNativeProbe]::GetForegroundWindow() -ne $sourceHandle -and [DateTime]::UtcNow -lt $until){Start-Sleep -Milliseconds 50}
                    if([ProtectionNativeProbe]::GetForegroundWindow() -ne $sourceHandle){throw 'Could not establish source-in-front precondition'}
                    $sourceElement=[System.Windows.Automation.AutomationElement]::FromHandle($sourceHandle)
                    if($request.allow){
                        $grant=$sourceElement.FindFirst([System.Windows.Automation.TreeScope]::Descendants,(New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::AutomationIdProperty,'AllowReaderForeground')))
                        ([System.Windows.Automation.InvokePattern]$grant.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)).Invoke()
                        $until=[DateTime]::UtcNow.AddSeconds(2)
                        while($grant.Current.Name -ne 'Permission granted' -and [DateTime]::UtcNow -lt $until){Start-Sleep -Milliseconds 40}
                        if($grant.Current.Name -ne 'Permission granted'){throw 'Source could not grant Reader foreground permission'}
                    }
                    $field=$sourceElement.FindFirst([System.Windows.Automation.TreeScope]::Descendants,(New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::AutomationIdProperty,'ReaderNumber')))
                    if($null -eq $field){throw 'Source ValuePattern field missing'}
                    ([System.Windows.Automation.ValuePattern]$field.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern)).SetValue([string]$request.text)
                    $pass=$true
                }else{
                    $until=[DateTime]::UtcNow.AddSeconds(6)
                    do{
                        $windows=@([ReaderWindowProbe]::Read([int[]]@($reader.Id))|Where-Object {$_.visible -and $_.title})
                        $main=@($windows|Where-Object {$_.title -eq 'Reader Data Viewer'})
                        $dialogs=@($windows|Where-Object {$_.title -ne 'Reader Data Viewer'})
                        $enabled=[System.Windows.Automation.AutomationElement]::FromHandle($mainHandle).Current.IsEnabled
                        $foreground=[ProtectionNativeProbe]::GetForegroundWindow().ToInt64()
                        $pass=$false
                        if($request.kind -eq 'main'){$pass=$enabled -and $dialogs.Count -eq 0 -and $foreground -eq $main[0].handle -and (-not [ProtectionNativeProbe]::IsIconic($mainHandle))}
                        if($request.kind -eq 'ready'){$pass=$enabled -and $dialogs.Count -eq 0}
                        if($request.kind -eq 'dialog'){$pass=(-not $enabled) -and $dialogs.Count -eq 1 -and [ProtectionNativeProbe]::OnScreen($dialogs[0])}
                        if($request.kind -eq 'away'){$pass=$foreground -eq $sourceHandle.ToInt64()}
                        if($request.foreground -and $request.kind -eq 'dialog'){$pass=$pass -and $foreground -eq $dialogs[0].handle}
                        if(-not $pass){Start-Sleep -Milliseconds 80}
                    }while(-not $pass -and [DateTime]::UtcNow -lt $until)
                }
                $reply=@{id=$request.id;kind=$request.kind;label=$request.label;pid=$reader.Id;utc=[DateTime]::UtcNow.ToString('o');pass=[bool]$pass;windows=$windows;foreground=[ProtectionNativeProbe]::GetForegroundWindow().ToInt64();sourceHandle=$sourceHandle.ToInt64();minimized=[ProtectionNativeProbe]::IsIconic($mainHandle);maximized=[ProtectionNativeProbe]::IsZoomed($mainHandle);topmost=[ProtectionNativeProbe]::Topmost($mainHandle);parentEnabled=[System.Windows.Automation.AutomationElement]::FromHandle($mainHandle).Current.IsEnabled}
                if($request.capture -and $pass){
                    $captureHandle=if($request.kind -eq 'dialog'){$dialogs[0].handle}else{$main[0].handle}
                    [ProtectionNativeProbe]::Capture($captureHandle,(Join-Path $Evidence ($request.label+'.png')))
                }
                [IO.File]::WriteAllText((Join-Path $Evidence 'native-reply.json'),($reply|ConvertTo-Json -Depth 6),$utf8)
                $handled=$request.id
            }
        }
        Start-Sleep -Milliseconds 60
    }
    if(-not $driver.HasExited){throw 'Native verification timed out'}
    $driver.WaitForExit()
    [IO.File]::WriteAllText((Join-Path $Evidence 'stdout.txt'),$stdout.Result,$utf8)
    [IO.File]::WriteAllText((Join-Path $Evidence 'stderr.txt'),$stderr.Result,$utf8)
    Write-Output $stdout.Result;Write-Output $stderr.Result
    if($driver.ExitCode -ne 0){throw ('Native verification failed: '+$driver.ExitCode)}
    if(-not $reader.WaitForExit(10000)){throw 'Reader did not close'}
    Write-Output ('Closed Reader PID '+$reader.Id+' exit='+$reader.ExitCode)
    $savedLines=[string[]]@();$savedStates=[string[]]@();$warning='';$savedProtection=$null
    [Rdv3Xlsx]::ReadProtected($ledger,$cfg.Data.Head,$cfg.Screen.Work.Column,[ref]$savedLines,[ref]$savedStates,[ref]$warning,[Rdv3Files]::StorageContract($cfg.Data,$cfg.Screen.Work),[Rdv3Files]::LegacyStorageContract($cfg.Data,$cfg.Screen.Work),[Rdv3BusinessDefinition]::Bound($cfg.Data),[ref]$savedProtection)
    if($AutomaticOnly){
        if($savedLines.Length -ne 100 -or $savedProtection.Deleted.Count -ne 0){throw 'Automatic condition changed ledger rows'}
        for($i=0;$i -lt $lines.Length;$i++){
            $expected=if($i -eq $paid[0] -or $i -eq $paid[1]){'TRUE'}else{'FALSE'}
            if($savedLines[$i] -ne $lines[$i] -or $savedStates[$i] -ne $expected){throw ('Automatic condition changed wrong record: '+$i)}
        }
        Write-Output 'PASS automatic condition actual XLSX readback: 100 full rows; only selected paid single and paid candidate saved TRUE; all other 98 FALSE'
    }else{
        if($savedLines.Length -ne 81 -or $savedProtection.Deleted.Count -ne 19 -or $savedLines[80] -ne $lines[0] -or $savedStates[80] -ne 'TRUE'){throw 'Restored XLSX content/state mismatch'}
        Write-Output 'PASS UI restore actual XLSX readback: 81 live, 19 archived, restored full original row and TRUE'
    }
}finally{
    if($source -and -not $source.HasExited){$source.CloseMainWindow()|Out-Null;$source.WaitForExit(3000)|Out-Null}
    foreach($process in @($driver,$reader,$source)){if($process -and -not $process.HasExited){Stop-Process -Id $process.Id -Force;$process.WaitForExit(3000)|Out-Null}}
    if($stdout){[IO.File]::WriteAllText((Join-Path $Evidence 'stdout.txt'),$stdout.Result,$utf8)}
    if($stderr){[IO.File]::WriteAllText((Join-Path $Evidence 'stderr.txt'),$stderr.Result,$utf8)}
    [IO.File]::WriteAllText((Join-Path $Evidence 'owned-pids.json'),($owned|ConvertTo-Json),$utf8)
    Write-Output ('Owned process cleanup: '+(($owned|ForEach-Object {'PID '+$_+' exists='+[bool](Get-Process -Id $_ -ErrorAction SilentlyContinue)}) -join ', '))
}
