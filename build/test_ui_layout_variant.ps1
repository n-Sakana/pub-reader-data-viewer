# Run one definition-driven layout extreme through the live WebView2 DOM.
# The caller supplies the preserved phase-18 settings fixture; product files
# are copied to a unique scratch directory and the window never activates.
[CmdletBinding()]
param(
    [string]$Root = '',
    [Parameter(Mandatory = $true)][string]$SettingsSource,
    [Parameter(Mandatory = $true)]
    [ValidateSet('emphasis-max', 'gap-zero')][string]$Scenario,
    [Parameter(Mandatory = $true)][string]$SingleKey,
    [int]$TimeoutSec = 90,
    [switch]$MutateKeyMax230
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
if (-not (Test-Path -LiteralPath $SettingsSource -PathType Leaf)) {
    throw "layout settings fixture is missing: $SettingsSource"
}
if ($SingleKey.Length -ne 8) {
    throw "the layout test requires an eight-character key: $SingleKey"
}

$scratchName = 'ui-dom-' + $Scenario
if ($MutateKeyMax230) { $scratchName += '-max230-mutant' }
$scratch = New-RdvTestDirectory -Root $Root -Name $scratchName
Copy-Item -LiteralPath (Join-Path $Root 'ReaderDataViewer.ps1') -Destination $scratch
Copy-Item -LiteralPath $SettingsSource -Destination (Join-Path $scratch 'settings.json')
foreach ($directory in 'src', 'lib', 'web', 'data') {
    Copy-Item -LiteralPath (Join-Path $Root $directory) -Destination $scratch -Recurse
}

# The test drives search itself and must not receive unrelated desktop input.
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

# Used only by the explicit mutation check. The shipped CSS is never edited;
# the exact regression is injected into this test-owned copy.
if ($MutateKeyMax230) {
    $cssPath = Join-Path $scratch 'web\app.css'
    $cssText = [IO.File]::ReadAllText($cssPath, [Text.Encoding]::UTF8)
    $needle = 'max-width:none;'
    if ([regex]::Matches($cssText, [regex]::Escape($needle)).Count -ne 1) {
        throw 'the key max-width mutation target is not unique'
    }
    $cssText = $cssText.Replace($needle, 'max-width:230px;')
    [IO.File]::WriteAllText(
        $cssPath,
        $cssText,
        (New-Object Text.UTF8Encoding($false)))
}

# Seed the same ledger shape as the default DOM suite so the requested search
# reaches a real record without adding a first-run confirmation to this check.
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
[Rdv3Xlsx]::Write(
    (Join-Path $scratch 'ReaderDataViewer-Ledger.xlsx'),
    $merge.Head,
    $cfg.Screen.Work.Column,
    $merge.Lines,
    $states,
    'layout-variant')

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
        (Join-Path $scratch 'ReaderDataViewer.ps1') + '"'
    $start.WorkingDirectory = $scratch
    $start.UseShellExecute = $false
    $start.CreateNoWindow = $true
    $start.EnvironmentVariables['RDV_WEBVIEW2_PROBE_OUTPUT'] = Join-Path $scratch 'surface.png'
    $start.EnvironmentVariables['RDV_WEBVIEW2_PROBE_KEEP_OPEN'] = '1'
    $start.EnvironmentVariables['RDV_WEBVIEW2_PROBE_MODAL_CAPTURE'] = Join-Path $scratch 'held-modal.png'
    $start.EnvironmentVariables['WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS'] =
        '--remote-debugging-port=' + $port
    $process = [Diagnostics.Process]::Start($start)

    & node.exe (Join-Path $Root 'build\test_ui_layout_variant.js') `
        $port `
        $SingleKey `
        $Scenario
    $result = $LASTEXITCODE

    if (-not $process.WaitForExit($TimeoutSec * 1000)) {
        Write-Output 'RESULT: FAIL (the variant app did not close gracefully)'
        $result = 1
    }
    elseif ($process.ExitCode -ne 0) {
        Write-Output ("RESULT: FAIL (variant app exit {0})" -f $process.ExitCode)
        $result = 1
    }
}
finally {
    if ($null -ne $process -and -not $process.HasExited) {
        Write-Output ("closing test-owned variant app process {0}" -f $process.Id)
        Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
    }
    Write-Output ("scratch: {0}" -f $scratch)
}

exit $result
