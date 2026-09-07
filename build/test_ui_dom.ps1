# Exercise the shipped WPF + WebView2 surface through its live DOM. The WPF
# window is created only in probe mode: off-screen, transparent, absent from
# the taskbar, and never activated.
[CmdletBinding()]
param(
    [string]$Root = '',
    [int]$TimeoutSec = 90
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
if ([string]::IsNullOrEmpty($Root)) {
    $Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
}

. (Join-Path $Root 'build\test_support.ps1')
Import-RdvProduct -Root $Root

if ($null -eq (Get-Command node.exe -ErrorAction SilentlyContinue)) {
    throw 'node.exe is required for the WebView2 DevTools DOM test'
}

$facts = @{}
$expectedPath = Join-Path $Root 'data\expected.txt'
if (-not (Test-Path -LiteralPath $expectedPath -PathType Leaf)) {
    throw "sample evidence is missing: $expectedPath"
}
foreach ($line in [IO.File]::ReadAllLines($expectedPath, [Text.Encoding]::UTF8)) {
    $equals = $line.IndexOf('=')
    if ($equals -gt 0) { $facts[$line.Substring(0, $equals)] = $line.Substring($equals + 1) }
}
foreach ($name in 'unmatchedA.firstkey', 'unmatchedA.firstidentity', 'cand5.firstkey') {
    if (-not $facts.ContainsKey($name) -or [string]::IsNullOrWhiteSpace($facts[$name])) {
        throw "expected.txt lacks $name"
    }
}

$scratch = New-RdvTestDirectory -Root $Root -Name 'ui-dom'
foreach ($file in 'settings.json') {
    Copy-Item -LiteralPath (Join-Path $Root $file) -Destination $scratch
}
foreach ($directory in 'src', 'lib', 'web', 'data') {
    Copy-Item -LiteralPath (Join-Path $Root $directory) -Destination $scratch -Recurse
}

# The test types its own values; an unrelated live Notepad must not feed the
# watcher. This changes only the unique scratch copy.
$settingsPath = Join-Path $scratch 'settings.json'
$settingsText = [IO.File]::ReadAllText($settingsPath, [Text.Encoding]::UTF8)
$settingsText = [regex]::Replace(
    $settingsText,
    '"targets": \[[\s\S]*?\n    \]',
    '"targets": []')
[IO.File]::WriteAllText(
    $settingsPath,
    $settingsText,
    (New-Object Text.UTF8Encoding($false)))

# Build a matching scratch ledger before launch. The DOM suite is about the
# stable READY surface; the guard suite separately exercises missing-ledger
# confirmation and creation.
$cfg = [Rdv3Config]::Load($settingsPath)
$dataDirectory = Join-Path $scratch 'data'
$heads = New-Object 'string[][]' $cfg.Data.Tables.Count
for ($index = 0; $index -lt $cfg.Data.Tables.Count; $index++) {
    $heads[$index] = [Rdv3Table]::ReadHead(
        (Join-Path $dataDirectory $cfg.Data.Tables[$index].File),
        $cfg.Data.Enc)
}
$cfg.Data.Bind($heads)
$merge = [Rdv3Ledger]::BuildFromCsv($cfg.Data, $dataDirectory)
$states = [Rdv3Ledger]::FreshStates(
    $merge.Lines.Length,
    $cfg.Screen.Work.InitialStored)
$ledgerPath = Join-Path $scratch $cfg.Ledger
[Rdv3Xlsx]::Write(
    $ledgerPath,
    $merge.Head,
    $cfg.Screen.Work.Column,
    $merge.Lines,
    $states,
    'dom-test')

$listener = New-Object Net.Sockets.TcpListener([Net.IPAddress]::Loopback, 0)
$listener.Start()
$port = ([Net.IPEndPoint]$listener.LocalEndpoint).Port
$listener.Stop()

$process = $null
$result = 1
try {
    $start = New-Object Diagnostics.ProcessStartInfo
    $start.FileName = 'powershell.exe'
    $start.Arguments = '-NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "' +
        (Join-Path $scratch 'src\ReaderDataViewer.ps1') + '"'
    $start.WorkingDirectory = $scratch
    $start.UseShellExecute = $false
    $start.CreateNoWindow = $true
    $start.EnvironmentVariables['RDV_WEBVIEW2_PROBE_OUTPUT'] = Join-Path $scratch 'surface.png'
    $start.EnvironmentVariables['RDV_WEBVIEW2_PROBE_KEEP_OPEN'] = '1'
    $start.EnvironmentVariables['RDV_WEBVIEW2_PROBE_MODAL_CAPTURE'] = Join-Path $scratch 'held-modal.png'
    $start.EnvironmentVariables['WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS'] =
        '--remote-debugging-port=' + $port
    $process = [Diagnostics.Process]::Start($start)

    & node.exe (Join-Path $Root 'build\test_ui_dom.js') `
        $port `
        $facts['unmatchedA.firstkey'] `
        $facts['unmatchedA.firstidentity'] `
        $facts['cand5.firstkey']
    $result = $LASTEXITCODE

    if (-not $process.WaitForExit($TimeoutSec * 1000)) {
        Write-Output 'RESULT: FAIL (the test-owned app did not close gracefully)'
        $result = 1
    }
    elseif ($process.ExitCode -ne 0) {
        Write-Output ("RESULT: FAIL (app exit {0})" -f $process.ExitCode)
        $result = 1
    }
}
finally {
    if ($null -ne $process -and -not $process.HasExited) {
        Write-Output ("closing test-owned app process {0}" -f $process.Id)
        Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
    }
    Remove-RdvTestDirectory -Path $scratch -Passed ($result -eq 0)
}

$variantScript = Join-Path $Root 'build\test_ui_layout_variant.ps1'

# The two layout extremes are derived from the shipped settings.json, into a
# directory this run owns. They used to be kept by hand under work\, which is
# scratch: one cleanup of that directory took the check with it. A screen that
# is built from a definition can always be re-derived from that definition.
#
#   emphasis-max  the three emphasis sizes at their ceiling (6..36 in
#                 Rdv3Screen), in a window wide enough to hold them
#   gap-zero      gap and padding at their floor (0), shipped window
#
# The expected pixels live in build\test_ui_layout_variant.js; if the shipped
# card definition drifts away from them, these checks fail, which is the point.
$fixtureRoot = New-RdvTestDirectory -Root $Root -Name 'ui-dom-layout-fixtures'
$settingsText = [IO.File]::ReadAllText(
    (Join-Path $Root 'settings.json'), [Text.Encoding]::UTF8)
$cardPattern = '(?m)^(\s*)("card":\s*\{[^\r\n]*\},)\s*$'
$cardMatches = [regex]::Matches($settingsText, $cardPattern)
if ($cardMatches.Count -ne 1) {
    throw 'settings.json must hold exactly one single-line screen.card definition'
}
$cardLine = $cardMatches[0].Groups[2].Value

function Set-RdvCardValues {
    param([string] $Card, [hashtable] $Values)

    foreach ($key in $Values.Keys) {
        $pattern = '("' + [regex]::Escape($key) + '":\s*)(\[[^\]]*\]|[^,}]+)'
        if ([regex]::Matches($Card, $pattern).Count -ne 1) {
            throw ("screen.card has no single {0} to set" -f $key)
        }
        $Card = [regex]::Replace($Card, $pattern, ('${1}' + $Values[$key]))
    }
    return $Card
}

function New-RdvLayoutFixture {
    param([string] $Name, [hashtable] $Values)

    $directory = Join-Path $fixtureRoot $Name
    New-Item -ItemType Directory -Path $directory | Out-Null
    $path = Join-Path $directory 'settings.json'
    $card = Set-RdvCardValues -Card $cardLine -Values $Values
    [IO.File]::WriteAllText(
        $path,
        [regex]::Replace($settingsText, $cardPattern, ('${1}' + $card)),
        (New-Object Text.UTF8Encoding($false)))
    return $path
}

$variantCases = @(
    @{
        Name = 'emphasis-max'
        Settings = New-RdvLayoutFixture -Name 'emphasis-max' -Values @{
            'width' = '1000'
            'startSize' = '[1000, 760]'
            'keyValueFontSize' = '36'
            'judgmentFontSize' = '36'
            'unsearchedFontSize' = '36'
        }
    },
    @{
        Name = 'gap-zero'
        Settings = New-RdvLayoutFixture -Name 'gap-zero' -Values @{
            'gap' = '0'
            'padding' = '[0]'
        }
    }
)
foreach ($variant in $variantCases) {
    Write-Output ''
    Write-Output ("=== layout variant: {0}" -f $variant.Name)
    & powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass `
        -File $variantScript `
        -Root $Root `
        -SettingsSource $variant.Settings `
        -Scenario $variant.Name `
        -SingleKey $facts['unmatchedA.firstkey'] `
        -TimeoutSec $TimeoutSec
    if ($LASTEXITCODE -ne 0) { $result = 1 }
}

Write-Output ''
Remove-RdvTestDirectory -Path $fixtureRoot -Passed ($result -eq 0)
if ($result -eq 0) {
    Write-Output '86 passed, 0 failed (70 default + 8 emphasis-max + 8 gap-zero)'
    Write-Output 'RESULT: PASS (full screen matrix)'
} else {
    Write-Output 'RESULT: FAIL (full screen matrix)'
}

exit $result
