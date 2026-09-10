# Run in the DEVELOPER distribution, not in a packaged application folder.
# Windows PowerShell 5.1 integration tests. Tests own a fresh temporary directory.
[CmdletBinding()]
param([switch]$Native)
Set-StrictMode -Version 2
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$builder = Join-Path $root 'tools/Build.ps1'
$packageName = 'ReaderDataViewer-json-layout'
if (Test-Path -LiteralPath (Join-Path $root 'src/Rdv3FixedScreen.cs')) { $packageName = 'ReaderDataViewer-fixed-layout' }
$hostExe = (Get-Process -Id $PID).Path
$temp = Join-Path ([IO.Path]::GetTempPath()) ('rdv-build-tests-' + [Guid]::NewGuid().ToString('N'))
[IO.Directory]::CreateDirectory($temp) | Out-Null
$results = New-Object System.Collections.Generic.List[object]
$utf8 = New-Object Text.UTF8Encoding($false)
function Assert([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
function Test([string]$Name, [scriptblock]$Body) {
    try { & $Body; $results.Add([pscustomobject]@{name=$Name;status='PASS'}); Write-Host ('PASS ' + $Name) }
    catch { $results.Add([pscustomobject]@{name=$Name;status='FAIL';detail=$_.Exception.ToString()}); Write-Host ('FAIL ' + $Name + ': ' + $_.Exception.Message) }
}
function Build([string[]]$Arguments, [int]$Expected = 0) {
    & $hostExe -NoProfile -ExecutionPolicy Bypass -File $builder @Arguments
    Assert ($LASTEXITCODE -eq $Expected) ('Unexpected build exit code: ' + $LASTEXITCODE)
}
function OnlyBuild([string]$Location) {
    $found = @(Get-ChildItem -LiteralPath $Location -Directory -Filter 'build-*')
    Assert ($found.Count -eq 1) 'Expected exactly one published build'
    return $found[0].FullName
}
try {
    Test 'powershell-parser' {
        foreach ($path in @($builder, (Join-Path $root 'src/ReaderDataViewer.ps1'), $PSCommandPath)) {
            $tokens=$null; $errors=$null
            [System.Management.Automation.Language.Parser]::ParseFile($path,[ref]$tokens,[ref]$errors) | Out-Null
            Assert ($errors.Count -eq 0) ('Parser errors in ' + $path + ': ' + ($errors | Out-String))
        }
    }
    $allOut = Join-Path $temp 'all'
    Test 'package-win98-with-explicit-skip' { Build @('package','-Theme','win98','-SkipValidation','-OutputRoot',$allOut) }
    Test 'metadata-hashes-data-and-zip' {
        $published = OnlyBuild $allOut
        $ids = @('win98')
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        foreach ($id in $ids) {
            $package = Join-Path $published $packageName
            $manifest = Get-Content -LiteralPath (Join-Path $published 'package-manifest.json') -Raw -Encoding UTF8 | ConvertFrom-Json
            Assert ($manifest.theme -eq $id -and $manifest.motion -eq 'off') 'Wrong Win98 metadata'
            Assert (-not (Test-Path -LiteralPath (Join-Path $package 'web/theme.json'))) 'Retired theme config was packaged'
            Assert ($manifest.validation.nativeCompile -eq 'not_run_explicit_skip') 'Skip must not be reported as a pass'
            $html = Get-Content -LiteralPath (Join-Path $package 'web/index.html') -Raw -Encoding UTF8
            Assert ($html -notmatch 'data-rdv-|themes\.css|theme-motion\.js') 'Retired theme asset was referenced'
            foreach ($file in $manifest.files) {
                $actual = Get-FileHash -LiteralPath (Join-Path $package $file.path) -Algorithm SHA256
                Assert ($actual.Hash.ToLowerInvariant() -eq $file.sha256) ('Hash mismatch: ' + $file.path)
            }
            Assert ((Get-FileHash -LiteralPath (Join-Path $package 'settings.json')).Hash -eq (Get-FileHash -LiteralPath (Join-Path $root 'configs/sample/settings.json')).Hash) 'Sample settings changed'
            Assert ((Get-FileHash -LiteralPath (Join-Path $package 'web/app.js')).Hash -eq (Get-FileHash -LiteralPath (Join-Path $root 'web/app.js')).Hash) 'Business browser logic changed'
            $dataFiles = @(Get-ChildItem -LiteralPath (Join-Path $package 'data') -File)
            Assert ($dataFiles.Count -eq 4 -and @($dataFiles | Where-Object { $_.Extension -eq '.csv' }).Count -eq 4) 'Expected four CSV five-record inputs only'
            foreach ($file in $dataFiles) {
                $expected = (Get-FileHash -LiteralPath (Join-Path $root ('samples/current/data/' + $file.Name))).Hash
                Assert ((Get-FileHash -LiteralPath $file.FullName).Hash -eq $expected) ('Initial data changed: ' + $file.Name)
            }
            Assert (-not (Test-Path -LiteralPath (Join-Path $package 'samples'))) 'Legacy 100-row samples were packaged'
            $sampleConfig = Get-Content -LiteralPath (Join-Path $package 'settings.json') -Raw -Encoding UTF8 | ConvertFrom-Json
            if (($sampleConfig.screen.PSObject.Properties.Name -contains 'sections')) { Assert (@($sampleConfig.screen.sections | Where-Object { $_.PSObject.Properties.Name -contains 'buttons' } | ForEach-Object { $_.buttons } | Where-Object { $_.action -eq 'restoreRecords' }).Count -eq 1) 'The restore action must remain available' }
            foreach ($name in @('tools/Migrate-Ledger.ps1','tools/ReaderDataViewer.cmd')) {
                Assert ((Get-FileHash -LiteralPath (Join-Path $package $name)).Hash -eq (Get-FileHash -LiteralPath (Join-Path $root $name)).Hash) ('Missing or changed delivery file: ' + $name)
            }
            Assert (-not (Test-Path -LiteralPath (Join-Path $package 'output'))) 'Runtime output was packaged; export creates it on demand'
            Assert (-not (Test-Path -LiteralPath (Join-Path $package 'configs/production'))) 'Production settings were packaged'
            Assert (-not (Test-Path -LiteralPath (Join-Path $package 'docs'))) 'Developer notes were packaged'
            Assert (-not (Test-Path -LiteralPath (Join-Path $package 'PACKAGE-README.txt'))) 'Duplicate startup documentation was packaged'
            $release = Get-Content -LiteralPath (Join-Path $package 'src/release.json') -Raw -Encoding UTF8 | ConvertFrom-Json
            Assert ($release.variant -in @('json-layout','fixed-layout')) 'Missing product variant'
            Assert ((@((Get-ChildItem -LiteralPath $package -File).Name | Sort-Object) -join ',') -eq 'LICENSE,ReaderDataViewer.vbs,README.md,settings.json') 'Unexpected files in the user entry folder'
            $zipPath = Join-Path $published ($packageName + '.zip')
            $archive = [IO.Compression.ZipFile]::OpenRead($zipPath)
            try {
                Assert ($archive.Entries.Count -gt 20) 'Empty or incomplete ZIP'
                foreach ($entry in $archive.Entries) {
                    Assert (-not $entry.FullName.Contains('\')) 'ZIP paths must use standard slashes'
                    $entryName = $entry.FullName.Replace('\','/')
                    Assert ($entryName.StartsWith($packageName + '/')) 'ZIP must contain one application root'
                    Assert (-not $entryName.Contains('../')) 'Unsafe ZIP entry'
                }
            } finally { $archive.Dispose() }
        }
    }
    Test 'repeat-build-never-overwrites' {
        $first = OnlyBuild $allOut
        $before = (Get-FileHash -LiteralPath (Join-Path $first 'build-summary.json')).Hash
        Build @('package','-Theme','win98','-SkipValidation','-OutputRoot',$allOut,'-Format','folder')
        Assert (@(Get-ChildItem -LiteralPath $allOut -Directory -Filter 'build-*').Count -eq 2) 'Second build overwrote first'
        Assert ((Get-FileHash -LiteralPath (Join-Path $first 'build-summary.json')).Hash -eq $before) 'First build changed'
    }
    Test 'no-samples-folder-only' {
        $output=Join-Path $temp 'subset'
        Build @('package','-Theme','win98','-Data','none','-Format','folder','-SkipValidation','-OutputRoot',$output)
        $published=OnlyBuild $output
        Assert (@(Get-ChildItem -LiteralPath $published -Directory).Count -eq 1) 'Incorrect package count'
        Assert (@(Get-ChildItem -LiteralPath $published -Filter '*.zip').Count -eq 0) 'ZIP created for folder-only output'
        foreach ($id in @('win98')) {
            $dir=Join-Path $published $packageName
            $theme=Get-Content -LiteralPath (Join-Path $published 'package-manifest.json') -Raw -Encoding UTF8 | ConvertFrom-Json
            Assert ($theme.motion -eq 'off') 'Win98 motion changed'
            Assert (@(Get-ChildItem -LiteralPath (Join-Path $dir 'data')).Count -eq 0) 'Data none ignored'
            Assert (-not (Test-Path -LiteralPath (Join-Path $dir 'samples'))) 'Samples included with Data none'
            Assert ((Get-FileHash -LiteralPath (Join-Path $dir 'settings.json')).Hash -eq (Get-FileHash -LiteralPath (Join-Path $root 'configs/sample/settings.json')).Hash) 'Non-sample settings changed'
        }
    }
    Test 'zip-only-output' {
        $output=Join-Path $temp 'zip'
        Build @('package','-Theme','win98','-Format','zip','-SkipValidation','-OutputRoot',$output)
        $published=OnlyBuild $output
        Assert (@(Get-ChildItem -LiteralPath $published -Directory).Count -eq 0) 'ZIP-only left an application directory'
        Assert ((Test-Path -LiteralPath (Join-Path $published ($packageName + '.zip')))) 'ZIP missing'
    }
    Test 'reject-unknown-theme' { Build @('package','-Theme','../invalid','-SkipValidation') 1 }
    Test 'default-win98' {
        $output = Join-Path $temp 'default'
        Build @('package','-SkipValidation','-Data','none','-Format','folder','-OutputRoot',$output)
        Assert (Test-Path -LiteralPath (Join-Path (OnlyBuild $output) $packageName)) 'Default variant package missing'
    }
    Test 'reject-retired-theme' { Build @('package','-Theme','apple','-SkipValidation') 1 }
    Test 'reject-all-plus-theme' { Build @('package','-All','-Theme','win98','-SkipValidation') 1 }
    Test 'reject-test-plus-skip' { Build @('package','-All','-RunTests','-SkipValidation') 1 }
    Test 'reject-output-in-input-tree' { Build @('package','-All','-SkipValidation','-OutputRoot',(Join-Path $root 'web/should-not-exist')) 1 }
    Test 'staging-cleaned' {
        Assert (@(Get-ChildItem -LiteralPath $temp -Recurse -Directory -Force | Where-Object { $_.Name -like '.rdv-stage-*' }).Count -eq 0) 'Staging directories remain'
    }
    if ($Native) {
        Test 'native-compile' { Build @('compile') }
        Test 'native-core-regression' { Build @('test') }
    }
} finally {
    $resolved = [IO.Path]::GetFullPath($temp)
    $owner = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd([IO.Path]::DirectorySeparatorChar)
    if (-not $resolved.StartsWith($owner + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase) -or [IO.Path]::GetFileName($resolved) -notlike 'rdv-build-tests-*') { throw ('Refusing unsafe cleanup: ' + $resolved) }
    if (Test-Path -LiteralPath $resolved) { Remove-Item -LiteralPath $resolved -Recurse -Force }
}
$failed = @($results | Where-Object { $_.status -eq 'FAIL' }).Count
$report = [ordered]@{scope='PowerShell packaging integration; native only with -Native'; powershell=$PSVersionTable.PSVersion.ToString(); utc=[DateTime]::UtcNow.ToString('o'); tests=@($results.ToArray()); failed=$failed; nativeRequested=[bool]$Native}
[IO.File]::WriteAllText((Join-Path $PSScriptRoot 'results/build-results.json'),($report | ConvertTo-Json -Depth 10),$utf8)
if ($failed) { exit 1 }
exit 0
