# Exercise the local-pending/send/exit boundary on the real WPF + WebView2
# product. UI operations and observations go through JavaScript and the live
# DOM; the probe window remains off-screen and is never activated.
[CmdletBinding()]
param(
    [string]$Root = '',
    [int]$ReadyTimeoutSec = 240,
    [int]$SaveTimeoutSec = 180
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
if ([string]::IsNullOrEmpty($Root)) {
    $Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
}

. (Join-Path $Root 'build\test_support.ps1')
Import-RdvProduct -Root $Root
Add-Type -AssemblyName System.IO.Compression.FileSystem

if ($null -eq (Get-Command node.exe -ErrorAction SilentlyContinue)) {
    throw 'node.exe is required for the WebView2 DevTools DOM test'
}

$sampleFacts = @{}
$sampleExpected = Join-Path $Root 'data-1k\expected.txt'
if (-not (Test-Path -LiteralPath $sampleExpected -PathType Leaf)) {
    throw "sample evidence is missing: $sampleExpected"
}
foreach ($line in [IO.File]::ReadAllLines($sampleExpected, [Text.Encoding]::UTF8)) {
    $equals = $line.IndexOf('=')
    if ($equals -gt 0) {
        $sampleFacts[$line.Substring(0, $equals)] = $line.Substring($equals + 1)
    }
}
foreach ($name in 'ledger.rows', 'cand1.firstkey', 'cand1.firstidentity') {
    if (-not $sampleFacts.ContainsKey($name) -or
        [string]::IsNullOrWhiteSpace($sampleFacts[$name])) {
        throw "expected.txt lacks $name"
    }
}
$ExpectedLedgerRows = [int]$sampleFacts['ledger.rows']
$TargetKey1 = $sampleFacts['cand1.firstkey']
$TargetIdentity = $sampleFacts['cand1.firstidentity']

$scratch = New-RdvTestDirectory -Root $Root -Name 'guard-run'
$appDirectory = Join-Path $scratch 'app'
New-Item -ItemType Directory -Path $appDirectory | Out-Null
foreach ($file in 'ReaderDataViewer.ps1', 'settings.json') {
    Copy-Item -LiteralPath (Join-Path $Root $file) -Destination $appDirectory
}
foreach ($directory in 'src', 'lib', 'web', 'data') {
    Copy-Item -LiteralPath (Join-Path $Root $directory) -Destination $appDirectory -Recurse
}

$settings = Join-Path $appDirectory 'settings.json'
$settingsText = [IO.File]::ReadAllText($settings, [Text.Encoding]::UTF8)
$settingsText = [regex]::Replace(
    $settingsText,
    '"targets": \[[\s\S]*?\n    \]',
    '"targets": []')
[IO.File]::WriteAllText(
    $settings,
    $settingsText,
    (New-Object Text.UTF8Encoding($false)))

$cfg = [Rdv3Config]::Load($settings)
$dataDirectory = Join-Path $appDirectory 'data'
$heads = New-Object 'string[][]' $cfg.Data.Tables.Count
for ($index = 0; $index -lt $cfg.Data.Tables.Count; $index++) {
    $heads[$index] = [Rdv3Table]::ReadHead(
        (Join-Path $dataDirectory $cfg.Data.Tables[$index].File),
        $cfg.Data.Enc)
}
$cfg.Data.Bind($heads)
$merge = [Rdv3Ledger]::BuildFromCsv($cfg.Data, $dataDirectory)
$initialStates = [Rdv3Ledger]::FreshStates(
    $merge.Lines.Length,
    $cfg.Screen.Work.InitialStored)
$ledger = Join-Path $appDirectory 'ReaderDataViewer-Ledger.xlsx'
[Rdv3Xlsx]::Write(
    $ledger,
    $merge.Head,
    $cfg.Screen.Work.Column,
    $merge.Lines,
    $initialStates,
    'guard-seed')

$log = Join-Path $appDirectory 'ReaderDataViewer.log'
$fakeLock = $ledger + '.lock'
$markerPath = $ledger + '.version'
$utf8 = New-Object Text.UTF8Encoding($false)
$outTsv = Join-Path $scratch 'checks.tsv'
[IO.File]::WriteAllText(
    $outTsv,
    "build`tcheck`tresult`tdetail`r`n",
    $utf8)

$checks = New-Object System.Collections.ArrayList
function Check([string]$build, [string]$name, [bool]$ok, [string]$detail) {
    $result = if ($ok) { 'PASS' } else { 'FAIL' }
    $clean = ($detail -replace "`t", ' ' -replace "`r?`n", ' ')
    [IO.File]::AppendAllText(
        $outTsv,
        ("{0}`t{1}`t{2}`t{3}`r`n" -f $build, $name, $result, $clean),
        $utf8)
    [void]$checks.Add([pscustomobject]@{
        Build = $build
        Name = $name
        Result = $result
        Detail = $detail
    })
    Write-Output ("  [{0}] {1}  {2}" -f $result, $name, $detail)
}

function Say([string]$text) {
    Write-Output ("{0}  {1}" -f (Get-Date -Format 'HH:mm:ss'), $text)
}

function Read-Log([string]$path, [Text.Encoding]$encoding) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return @() }
    for ($attempt = 0; $attempt -lt 20; $attempt++) {
        try { return @([IO.File]::ReadAllLines($path, $encoding)) }
        catch { Start-Sleep -Milliseconds 100 }
    }
    return @()
}

function Wait-Log(
    [string]$path,
    [Text.Encoding]$encoding,
    [string]$pattern,
    [int]$from,
    [int]$seconds,
    [int]$pollMilliseconds) {
    $started = Get-Date
    while (((Get-Date) - $started).TotalSeconds -lt $seconds) {
        $all = Read-Log $path $encoding
        for ($index = $from; $index -lt $all.Count; $index++) {
            if ($all[$index] -match $pattern) { return , @($index, $all[$index]) }
        }
        Start-Sleep -Milliseconds $pollMilliseconds
    }
    return $null
}

function Count-TrueRows([string]$path) {
    $archive = [IO.Compression.ZipFile]::OpenRead($path)
    try {
        $sheet = $null
        foreach ($entry in $archive.Entries) {
            if ($entry.FullName -like 'xl/worksheets/*.xml' -and
                ($null -eq $sheet -or $entry.Length -gt $sheet.Length)) {
                $sheet = $entry
            }
        }
        if ($null -eq $sheet) { return -1 }
        $reader = New-Object IO.StreamReader($sheet.Open(), [Text.Encoding]::UTF8)
        try {
            $count = 0
            $buffer = New-Object char[] 262144
            $carry = ''
            while (($read = $reader.Read($buffer, 0, $buffer.Length)) -gt 0) {
                $text = $carry + (New-Object string($buffer, 0, $read))
                $count += ([regex]::Matches(
                    $text,
                    '<row[^>]*><c[^>]*><is><t>TRUE</t>')).Count
                if ($text.Length -gt 64) { $carry = $text.Substring($text.Length - 64) }
                else { $carry = $text }
            }
            return $count
        }
        finally { $reader.Dispose() }
    }
    finally { $archive.Dispose() }
}

function Write-TestMarker(
    [string]$path,
    [long]$version,
    [string]$kind,
    [int]$done,
    [int]$todo) {
    $host64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes('REMOTE-HOST'))
    $user64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes('remote-user'))
    $writer64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes('remote-writer'))
    $text = "RDV-MARKER-1`t$version`t$host64`t$user64`t$writer64`t$([DateTime]::UtcNow.Ticks)`t$ExpectedLedgerRows`t$kind`t$done`t$todo`r`n"
    $temporary = $path + '.test-tmp'
    [IO.File]::WriteAllText($temporary, $text, $utf8)
    [IO.File]::Replace($temporary, $path, [NullString]::Value)
}

function Write-ChangedLedger(
    [string]$ledgerPath,
    [string]$settingsPath,
    [string]$dataPath,
    [string]$identity) {
    $site = [Rdv3Config]::Load($settingsPath)
    $siteHeads = New-Object 'string[][]' $site.Data.Tables.Count
    for ($index = 0; $index -lt $site.Data.Tables.Count; $index++) {
        $siteHeads[$index] = [Rdv3Table]::ReadHead(
            (Join-Path $dataPath $site.Data.Tables[$index].File),
            $site.Data.Enc)
    }
    $site.Data.Bind($siteHeads)
    $lines = $null
    $states = $null
    [Rdv3Xlsx]::Read(
        $ledgerPath,
        $site.Data.Head,
        $site.Screen.Work.Column,
        [ref]$lines,
        [ref]$states)
    $row = -1
    for ($index = 0; $index -lt $lines.Length; $index++) {
        if ([Rdv3Ledger]::FieldOf($lines[$index], $site.Data.IdentityCol) -eq $identity) {
            $row = $index
            break
        }
    }
    if ($row -lt 0) { throw "identity not found in scratch ledger: $identity" }
    $cells = [Rdv3Ledger]::SplitLine($lines[$row])
    $cells[2] = $cells[2] + '-REMOTE'
    $lines[$row] = [string]::Join("`t", $cells)
    $states[$row] = $site.Screen.Work.InitialStored

    $remote = New-Object -TypeName Rdv3SharedFiles -ArgumentList @(
        $ledgerPath,
        'REMOTE-HOST',
        'remote-user',
        'remote-writer')
    $owner = $null
    $lease = $remote.TryAcquire([ref]$owner)
    if ($null -eq $lease) { throw 'the remote test writer could not acquire the scratch lock' }
    try {
        [Rdv3Xlsx]::Write(
            $ledgerPath,
            $site.Data.Head,
            $site.Screen.Work.Column,
            $lines,
            $states,
            'remote-update')
        return $remote.WriteMarker('update', $lines.Length, 0, 0)
    }
    finally { $lease.Release() }
}

$driver = Join-Path $Root 'build\webview2_cdp.js'
$port = 0
function Invoke-Cdp([string[]]$CdpArguments) {
    $output = @(& node.exe $driver $port @CdpArguments 2>&1)
    $exitCode = $LASTEXITCODE
    $text = ($output | ForEach-Object { $_.ToString() }) -join [Environment]::NewLine
    if ($exitCode -ne 0) { throw "WebView2 DOM command failed ($exitCode): $text" }
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    return ($text | ConvertFrom-Json)
}

function Encode-Dom([string]$expression) {
    return [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($expression))
}

function Invoke-Dom([string]$expression) {
    return Invoke-Cdp @('eval', (Encode-Dom $expression))
}

function Wait-Dom([string]$expression, [int]$milliseconds) {
    return Invoke-Cdp @('wait', (Encode-Dom $expression), $milliseconds.ToString())
}

function Press-Dom([string]$key, [int]$modifiers = 0) {
    return Invoke-Cdp @('press', $key, $modifiers.ToString())
}

function Search-Dom([string]$key) {
    $quoted = $key | ConvertTo-Json -Compress
    [void](Invoke-Dom ("(() => { const input = document.querySelector('#input'); " +
        "input.textContent = $quoted; " +
        "input.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText' })); " +
        "input.focus(); document.querySelector('#b-search').click(); return true; })()"))
}

function Click-Dom([string]$selector) {
    $quoted = $selector | ConvertTo-Json -Compress
    [void](Invoke-Dom ("(() => { const node = document.querySelector($quoted); " +
        "node.focus(); node.click(); return true; })()"))
}

function Test-Dom([string]$expression) {
    try { return [bool](Invoke-Dom $expression) }
    catch { return $false }
}

$listener = New-Object Net.Sockets.TcpListener([Net.IPAddress]::Loopback, 0)
$listener.Start()
$port = ([Net.IPEndPoint]$listener.LocalEndpoint).Port
$listener.Stop()

$process = $null
$pendingFile = ''
$harnessError = $null
$build = 'wpf-webview2'

Write-Output ''
Write-Output '=== WPF + WebView2 build'
Say ("target key1 = " + $TargetKey1)

try {
    $start = New-Object Diagnostics.ProcessStartInfo
    $start.FileName = 'powershell.exe'
    $start.Arguments = '-NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "' +
        (Join-Path $appDirectory 'ReaderDataViewer.ps1') + '"'
    $start.WorkingDirectory = $appDirectory
    $start.UseShellExecute = $false
    $start.CreateNoWindow = $true
    $start.EnvironmentVariables['RDV_WEBVIEW2_PROBE_OUTPUT'] = Join-Path $scratch 'surface.png'
    $start.EnvironmentVariables['RDV_WEBVIEW2_PROBE_KEEP_OPEN'] = '1'
    $start.EnvironmentVariables['RDV_WEBVIEW2_PROBE_MODAL_CAPTURE'] = Join-Path $scratch 'held-modal.png'
    $start.EnvironmentVariables['WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS'] =
        '--remote-debugging-port=' + $port
    $process = [Diagnostics.Process]::Start($start)

    [void](Invoke-Cdp @('install', '60000'))
    [void](Wait-Dom "document.querySelector('.stage.runtime #b-search[aria-disabled=false]') !== null" ($ReadyTimeoutSec * 1000))
    $initialReady = Wait-Log $log $utf8 "`tdecision`tready " 0 $ReadyTimeoutSec 100
    if ($null -eq $initialReady) { throw 'the app never reached its seeded READY state' }

    $pendingLine = Wait-Log $log $utf8 "`tpending`tpath=.* count=" 0 10 100
    if ($null -ne $pendingLine -and
        $pendingLine[1] -match "`tpending`tpath=(.*) count=[0-9]+$") {
        $pendingFile = $Matches[1]
    }

    # Exercise the missing-ledger branch after the DOM hook is installed. It
    # is the same StartCheck/EndCheck/StartApply path as first launch, without
    # racing the sub-100 ms input merge against DevTools attachment.
    $scratchPrefix = [IO.Path]::GetFullPath($scratch).TrimEnd('\') + '\'
    $ledgerFull = [IO.Path]::GetFullPath($ledger)
    if (-not $ledgerFull.StartsWith($scratchPrefix, [StringComparison]::OrdinalIgnoreCase)) {
        throw "refusing to remove a ledger outside scratch: $ledgerFull"
    }
    $refreshFrom = (Read-Log $log $utf8).Count
    [IO.File]::Delete($ledgerFull)
    [void](Invoke-Dom "window.chrome.webview.postMessage({ type: 'action', name: 'refreshLedger', job: '', key: '' }); true")
    [void](Wait-Dom "document.querySelector('#v-send').classList.contains('show')" 30000)
    $createDialog = Invoke-Dom "(() => { const veil = document.querySelector('#v-send'); return { label: veil.querySelector('.dlg').getAttribute('aria-label'), body: veil.querySelector('.modal-message').innerText, buttons: Array.from(veil.querySelectorAll('.foot .btn')).map(node => node.textContent) }; })()"
    Click-Dom '#v-send .foot [data-modal-default=true]'
    $createApproved = Wait-Log $log $utf8 "ledger missing; create approved" $refreshFrom 30 100
    $ready = Wait-Log $log $utf8 "`tdecision`tready rows=$ExpectedLedgerRows " $refreshFrom $ReadyTimeoutSec 100
    $readyOk = $null -ne $createApproved -and $null -ne $ready -and
        (Test-Path -LiteralPath $ledger -PathType Leaf) -and
        $createDialog.body -like '*統合台帳がありません*' -and
        $createDialog.buttons.Count -eq 2
    Check $build 'ready' $readyOk $(if ($null -ne $ready) {
        'missing-ledger DOM confirmation approved; ' + $ready[1]
    } else { 'missing-ledger branch did not return to READY' })
    Check $build 'pending_path' (-not [string]::IsNullOrEmpty($pendingFile)) $pendingFile
    $windowDom = Invoke-Dom "(() => ({ host: location.hostname, runtime: document.querySelector('.stage').classList.contains('runtime'), windows: document.querySelectorAll('.stage > .win').length }))()"
    Check $build 'window' (-not $process.HasExited -and
        $windowDom.host -eq 'reader-data-viewer.local' -and
        $windowDom.runtime -and $windowDom.windows -eq 1) ("pid=" + $process.Id + '; dom=' + ($windowDom | ConvertTo-Json -Compress))

    $searchFrom = (Read-Log $log $utf8).Count
    Search-Dom $TargetKey1
    $hit = Wait-Log $log $utf8 ("`tsearch`tkey=" + $TargetKey1 + " ") $searchFrom 30 100
    [void](Wait-Dom ("Array.from(document.querySelectorAll('.fld')).some(node => node.textContent.trim() === " +
        ($TargetIdentity | ConvertTo-Json -Compress) + ")") 30000)
    Check $build 'search' ($null -ne $hit) $(if ($null -ne $hit) { $hit[1] } else { 'no search line' })

    $stateFrom = (Read-Log $log $utf8).Count
    Click-Dom '#b-work'
    [void](Wait-Dom "document.querySelector('#v-send').classList.contains('show')" 30000)
    $stateConfirm = Invoke-Dom "(() => { const veil = document.querySelector('#v-send'); return { label: veil.querySelector('.dlg').getAttribute('aria-label'), body: veil.querySelector('.modal-message').innerText, yes: veil.querySelector('[data-modal-default=true]').textContent }; })()"
    Click-Dom '#v-send [data-modal-default=true]'
    Check $build 'confirm' ($stateConfirm.label -eq '処理済の確認' -and
        $stateConfirm.body -like ("*" + $TargetIdentity + "*") -and
        $stateConfirm.yes -like 'はい*') ('DOM confirmation: ' + ($stateConfirm | ConvertTo-Json -Compress))

    $started = Wait-Log $log $utf8 'save started .*exit held' $stateFrom 20 15
    Check $build 'save_started' ($null -ne $started) $(if ($null -ne $started) { $started[1] } else { 'no save-started line' })
    $done = Wait-Log $log $utf8 "`tstate`tkey2=.*value=TRUE" $stateFrom $SaveTimeoutSec 100
    Check $build 'save_completed' ($null -ne $done) $(if ($null -ne $done) { $done[1] } else { 'no state line' })
    $released = Wait-Log $log $utf8 'write decided .*exit released' $stateFrom 20 100
    Check $build 'exit_released' ($null -ne $released) $(if ($null -ne $released) { $released[1] } else { 'no release line' })
    $pendingOne = Invoke-Dom "document.querySelector('#sn').textContent"
    Check $build 'pending_count_one' ($pendingOne -eq '未送信 1 件') $pendingOne
    Check $build 'shared_untouched_before_send' ((Count-TrueRows $ledger) -eq 0) 'processed=TRUE rows before send: expected 0'
    $pendingLength = if (Test-Path -LiteralPath $pendingFile -PathType Leaf) {
        (Get-Item -LiteralPath $pendingFile).Length
    } else { -1 }
    Check $build 'small_local_pending' ($pendingLength -ge 0 -and $pendingLength -lt 512) ("bytes=" + $pendingLength)

    $host64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes('TEST-HOST'))
    $user64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes('lock-owner'))
    $lockText = "RDV-LOCK-1`t$host64`t$user64`t$([DateTime]::UtcNow.Ticks)`r`n"
    [IO.File]::WriteAllText($fakeLock, $lockText, $utf8)
    [IO.File]::SetLastWriteTimeUtc($fakeLock, [DateTime]::UtcNow.AddMinutes(-11))

    $sendFrom = (Read-Log $log $utf8).Count
    Click-Dom '#b-send'
    [void](Wait-Dom "document.querySelector('#v-send').classList.contains('show')" 30000)
    $sendDialog = Invoke-Dom "(() => { const veil = document.querySelector('#v-send'); return { label: veil.querySelector('.dlg').getAttribute('aria-label'), body: veil.querySelector('.modal-message').innerText, yes: veil.querySelector('[data-modal-default=true]').textContent }; })()"
    Check $build 'send_confirm_body' ($sendDialog.label -eq '送信' -and
        $sendDialog.body -like '*未送信の 1 件を送信します。よろしいですか*') $sendDialog.body
    Check $build 'send_confirm_effect' ($sendDialog.body -like
        '*送信した行は未送信から外れます。取り込んだデータは書き換えません。*') $sendDialog.body
    Click-Dom '#v-send [data-modal-default=true]'
    Check $build 'send_confirm' ($sendDialog.yes -like 'はい*') 'answered はい through the DOM'

    $stale = Wait-Log $log $utf8 "`tlock`tremoved stale lock age_ms=" $sendFrom 20 100
    Check $build 'stale_lock_removed' ($null -ne $stale) $(if ($null -ne $stale) { $stale[1] } else { 'no stale-removal line' })
    $staleAge = 0
    if ($null -ne $stale -and $stale[1] -match 'age_ms=([0-9]+)') {
        $staleAge = [long]$Matches[1]
    }
    Check $build 'stale_lock_threshold' ($staleAge -ge 600000) ("age_ms=" + $staleAge)
    $acquired = Wait-Log $log $utf8 "`tlock`tacquired " $sendFrom 20 100
    Check $build 'lock_reacquired_after_stale' ($null -ne $acquired) $(if ($null -ne $acquired) { $acquired[1] } else { 'no acquired line' })
    $aliveDom = Test-Dom "document.querySelector('.stage.runtime > .win') !== null"
    Check $build 'window_alive_after_stale' (-not $process.HasExited -and $aliveDom) ("pid=" + $process.Id + '; live DOM=' + $aliveDom)

    $sent = Wait-Log $log $utf8 "`tmarker`tversion=.* kind=send" $sendFrom $SaveTimeoutSec 100
    Check $build 'send_completed' ($null -ne $sent) $(if ($null -ne $sent) { $sent[1] } else { 'no marker line' })
    $sendReady = Wait-Log $log $utf8 "`tshared`tready .* pending=0" $sendFrom 30 100
    Check $build 'pending_cleared' ($null -ne $sendReady) $(if ($null -ne $sendReady) { $sendReady[1] } else { 'pending did not clear' })
    [void](Wait-Dom "document.querySelector('#sn').textContent === '未送信 0 件'" 30000)
    $pendingZero = Invoke-Dom "document.querySelector('#sn').textContent"

    $repeatFrom = (Read-Log $log $utf8).Count
    Search-Dom $TargetKey1
    $repeatHit = Wait-Log $log $utf8 ("`tsearch`tkey=" + $TargetKey1 + " ") $repeatFrom 30 100
    [void](Wait-Dom "document.querySelector('#b-work').getAttribute('aria-disabled') === 'false'" 30000)
    $stateButton = Invoke-Dom "(() => { const node = document.querySelector('#b-work'); return { text: node.textContent, disabled: node.getAttribute('aria-disabled'), tab: node.tabIndex }; })()"
    Check $build 'pending_count_zero' ($pendingZero -eq '未送信 0 件') $pendingZero
    Check $build 'sent_row_not_frozen' ($null -ne $repeatHit -and
        $stateButton.text -like '処理済*' -and
        $stateButton.disabled -eq 'false' -and $stateButton.tab -eq 0) ($stateButton | ConvertTo-Json -Compress)
    Check $build 'shared_after_send' ((Count-TrueRows $ledger) -eq 1) 'processed=TRUE rows after send: expected 1'
    $markerLines = if (Test-Path -LiteralPath $markerPath -PathType Leaf) {
        @([IO.File]::ReadAllLines($markerPath, $utf8))
    } else { @() }
    Check $build 'marker_written' ($markerLines.Count -eq 1) ("marker lines=" + $markerLines.Count)
    Check $build 'lock_released' (-not (Test-Path -LiteralPath $fakeLock)) 'no lock file remains'

    $markerReader = New-Object -TypeName Rdv3SharedFiles -ArgumentList @(
        $ledger,
        'REMOTE-HOST',
        'remote-user',
        'remote-reader')
    $observedMarker = $markerReader.ReadMarker()
    if ($null -eq $observedMarker) { throw 'the scratch ledger has no marker after send' }
    $remoteFirstVersion = [long]($observedMarker.Version + 1)
    $remoteSecondVersion = [long]($observedMarker.Version + 2)
    $remoteUpdateVersion = [long]($observedMarker.Version + 3)
    $remoteFrom = (Read-Log $log $utf8).Count
    Write-TestMarker $markerPath $remoteFirstVersion 'send' 12 3
    $remoteFirst = Wait-Log $log $utf8 "`tnotice`ttarget=status text=remote-user が 12 件を処理済、3 件を未処理にしました" $remoteFrom 15 100
    Write-TestMarker $markerPath $remoteSecondVersion 'send' 7 4
    $remoteSecond = Wait-Log $log $utf8 "`tnotice`ttarget=status text=remote-user が 7 件を処理済、4 件を未処理にしました" $remoteFrom 15 100
    $remoteStatus = Invoke-Dom "(() => { const notice = document.querySelector('.sp.notice'); return { text: notice ? notice.textContent : '', dialogs: document.querySelectorAll('.veil.show').length }; })()"
    Check $build 'remote_send_status' ($null -ne $remoteFirst -and
        $null -ne $remoteSecond -and
        $remoteStatus.text -eq 'remote-user が 7 件を処理済、4 件を未処理にしました' -and
        $remoteStatus.dialogs -eq 0) ('DOM status=' + ($remoteStatus | ConvertTo-Json -Compress))
    $remoteReload = Wait-Log $log $utf8 ("`treload`tversion=" + $remoteSecondVersion + " .*rows=" + $ExpectedLedgerRows) $remoteFrom 30 100
    Check $build 'remote_send_reload' ($null -ne $remoteReload) $(if ($null -ne $remoteReload) { $remoteReload[1] } else { 'no reload line' })
    $remoteReady = Wait-Log $log $utf8 ("`tshared`tready .*note=marker-" + $remoteSecondVersion) $remoteFrom 30 100
    Check $build 'remote_send_ready' ($null -ne $remoteReady) $(if ($null -ne $remoteReady) { $remoteReady[1] } else { 'latest remote marker was not adopted' })

    $updateFrom = (Read-Log $log $utf8).Count
    $changedMarker = Write-ChangedLedger $ledger $settings $dataDirectory $TargetIdentity
    Check $build 'remote_update_written' ($changedMarker.Version -eq $remoteUpdateVersion -and
        (Count-TrueRows $ledger) -eq 0) ('version ' + $changedMarker.Version + ', processed=TRUE rows 0')
    [void](Wait-Dom "document.querySelector('#v-shared').classList.contains('show')" 30000)
    $updateDialog = Invoke-Dom "(() => { const veil = document.querySelector('#v-shared'); const table = veil.querySelector('.lv table'); return { label: veil.querySelector('.dlg').getAttribute('aria-label'), hint: veil.querySelector('.hint').textContent, tables: veil.querySelectorAll('.lv table').length, rows: table ? table.querySelectorAll('tbody tr').length : 0, columns: table ? table.querySelectorAll('thead th').length : 0 }; })()"
    Check $build 'remote_update_dialog' ($updateDialog.label -eq '台帳の更新' -and
        $updateDialog.hint -like '*台帳が更新されました。切り替えますか*') ($updateDialog | ConvertTo-Json -Compress)
    Check $build 'remote_update_reset_count' ($updateDialog.hint -like
        '*中身が変わったため未処理に戻ったレコード: 1 件*') $updateDialog.hint
    Check $build 'remote_update_reset_list' ($updateDialog.tables -eq 1 -and
        $updateDialog.rows -eq 1 -and $updateDialog.columns -eq 10) ($updateDialog | ConvertTo-Json -Compress)
    Click-Dom '#v-shared [data-modal-default=true]'
    $updateReady = Wait-Log $log $utf8 ("`tshared`tready .*note=marker-" + $remoteUpdateVersion) $updateFrom 30 100
    Check $build 'remote_update_switch' ($null -ne $updateReady) $(if ($null -ne $updateReady) { $updateReady[1] } else { 'no latest-marker ready line' })

    $closeFrom = (Read-Log $log $utf8).Count
    Search-Dom $TargetKey1
    $closeHit = Wait-Log $log $utf8 ("`tsearch`tkey=" + $TargetKey1 + " ") $closeFrom 30 100
    if ($null -eq $closeHit) { throw 'the close-while-waiting setup search did not finish' }
    Click-Dom '#b-work'
    [void](Wait-Dom "document.querySelector('#v-send').classList.contains('show')" 30000)
    Click-Dom '#v-send [data-modal-default=true]'
    $closeSaved = Wait-Log $log $utf8 "`tstate`tkey2=.*value=TRUE" $closeFrom $SaveTimeoutSec 100
    if ($null -eq $closeSaved) { throw 'the close-while-waiting setup state did not save' }

    $lockText = "RDV-LOCK-1`t$host64`t$user64`t$([DateTime]::UtcNow.Ticks)`r`n"
    [IO.File]::WriteAllText($fakeLock, $lockText, $utf8)
    $closeSendFrom = (Read-Log $log $utf8).Count
    Click-Dom '#b-send'
    [void](Wait-Dom "document.querySelector('#v-send').classList.contains('show')" 30000)
    Click-Dom '#v-send [data-modal-default=true]'
    $closeWaiting = Wait-Log $log $utf8 "`tlock`twaiting owner=lock-owner host=TEST-HOST" $closeSendFrom 20 100
    if ($null -eq $closeWaiting) { throw 'the send did not wait behind the live lock' }
    Click-Dom '.stage > .win > .tb .cl'
    $closed = $process.WaitForExit(10000)
    $closingLog = Wait-Log $log $utf8 "`texit`tclosing" $closeSendFrom 5 100
    Check $build 'closes_while_lock_waiting' ($closed -and $null -ne $closingLog) ('closed=' + $closed + '; wait=' + $closeWaiting[1] + '; closing=' + $(if ($null -ne $closingLog) { $closingLog[1] } else { 'missing' }))

    $trueRows = Count-TrueRows $ledger
    Check $build 'changed_record_reset' ($trueRows -eq 0) ("processed=TRUE rows after the changed-record update: " + $trueRows)
}
catch {
    $harnessError = $_.Exception.Message + ' @ ' + $_.InvocationInfo.ScriptLineNumber +
        ' | ' + ($_.ScriptStackTrace -replace "`r?`n", ' <- ')
    Write-Output ("HARNESS ERROR: " + $harnessError)
}
finally {
    if ($null -ne $process -and -not $process.HasExited) {
        try {
            [void](Invoke-Dom "window.__rdvTestNativePost({ type: 'window', command: 'close' }); true")
            [void]$process.WaitForExit(3000)
        }
        catch { }
    }
    if ($null -ne $process -and -not $process.HasExited) {
        Say ("closing test-owned app process " + $process.Id)
        Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
    }
    if (Test-Path -LiteralPath $fakeLock -PathType Leaf) {
        $lockFull = [IO.Path]::GetFullPath($fakeLock)
        $scratchFull = [IO.Path]::GetFullPath($scratch).TrimEnd('\') + '\'
        if ($lockFull.StartsWith($scratchFull, [StringComparison]::OrdinalIgnoreCase)) {
            Remove-Item -LiteralPath $lockFull -Force
        }
    }
    if (-not [string]::IsNullOrEmpty($pendingFile) -and
        (Test-Path -LiteralPath $pendingFile -PathType Leaf)) {
        $pendingFull = [IO.Path]::GetFullPath($pendingFile)
        $pendingRoot = Join-Path (
            [Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)) `
            'ReaderDataViewer'
        $pendingPrefix = [IO.Path]::GetFullPath($pendingRoot).TrimEnd('\') + '\'
        if ($pendingFull.StartsWith($pendingPrefix, [StringComparison]::OrdinalIgnoreCase) -and
            [IO.Path]::GetFileName($pendingFull) -like 'pending-*.dat') {
            Remove-Item -LiteralPath $pendingFull -Force
        }
    }
}

Write-Output ''
$pass = @($checks | Where-Object { $_.Result -eq 'PASS' }).Count
$fail = @($checks | Where-Object { $_.Result -eq 'FAIL' }).Count
Write-Output ("=== {0} checks: {1} PASS, {2} FAIL   -> {3}" -f
    $checks.Count, $pass, $fail, $outTsv)
foreach ($check in $checks) {
    if ($check.Result -eq 'FAIL') {
        Write-Output ("  FAIL {0}/{1}: {2}" -f $check.Build, $check.Name, $check.Detail)
    }
}
if ($null -ne $harnessError) { Write-Output ("  HARNESS: " + $harnessError) }
if ($checks.Count -ne 35) {
    Write-Output ("  FAIL assertion count: expected 35, actual {0}" -f $checks.Count)
    $fail++
}
Write-Output ("scratch: {0}" -f $scratch)
if ($fail -gt 0) { exit 1 }
Write-Output 'RESULT: PASS'
