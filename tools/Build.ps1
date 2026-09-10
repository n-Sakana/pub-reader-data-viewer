# Windows PowerShell 5.1 compatible; no npm, SDK, downloaded UI kit or admin needed.
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [ValidateSet('menu', 'compile', 'test', 'package', 'list', 'help')]
    [string]$Command = 'menu',
    [Alias('Themes')][string]$Theme = '',
    [switch]$All,
    [ValidateSet('both', 'folder', 'zip')][string]$Format = 'both',
    [ValidateSet('sample', 'none')][string]$Data = 'sample',
    [string]$OutputRoot = '',
    [string]$SampleRoot = '',
    [switch]$RunTests,
    [switch]$SkipValidation
)
Set-StrictMode -Version 2
$ErrorActionPreference = 'Stop'
Import-Module Microsoft.PowerShell.Utility
$script:Root = Split-Path -Parent $PSScriptRoot
$script:Utf8 = New-Object System.Text.UTF8Encoding($false)

function Write-Utf8([string]$Path, [string]$Text) {
    [IO.File]::WriteAllText($Path, $Text, $script:Utf8)
}
function Write-Json([string]$Path, $Object) {
    Write-Utf8 $Path (($Object | ConvertTo-Json -Depth 10) + [Environment]::NewLine)
}
function Show-Usage {
    Write-Host (@(
        'Reader Data Viewer - build'
        ''
        '  build.bat                         Choose format, sample data and tests'
        '  build.bat compile                 Compile the C# application'
        '  build.bat test                    Run isolated C# regression tests'
        '  build.bat package                  Create a new folder and ZIP'
        '  build.bat package -Format zip -Data none -OutputRoot "C:\RDV releases"'
        ''
        'Options: -Format both|folder|zip, -Data sample|none, -OutputRoot PATH'
        '         -SampleRoot PATH (local terminology bundle), -RunTests, -SkipValidation'
        'The branch selects the product. -Theme win98 and -All are compatibility aliases.'
        'Each build creates a NEW folder. Existing packages and input files stay intact.'
        'This is a source-at-startup application, not a standalone EXE.'
        'See docs/design-build.md for Japanese instructions.'
    ) -join [Environment]::NewLine)
}
function Invoke-NativeCheck([string]$Mode) {
    if ($env:OS -ne 'Windows_NT') {
        throw 'Windows C#/WPF verification requires Windows PowerShell 5.1 on Windows. For packaging tests ONLY, use -SkipValidation.'
    }
    $exe = Join-Path $env:SystemRoot 'System32/WindowsPowerShell/v1.0/powershell.exe'
    if (Test-Path -LiteralPath (Join-Path $env:SystemRoot 'Sysnative/WindowsPowerShell/v1.0/powershell.exe')) {
        $exe = Join-Path $env:SystemRoot 'Sysnative/WindowsPowerShell/v1.0/powershell.exe'
    }
    $flag = '-CompileOnly'
    if ($Mode -eq 'test') { $flag = '-TestCore' }
    & $exe -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File (Join-Path $script:Root 'src/ReaderDataViewer.ps1') $flag
    if ($LASTEXITCODE -ne 0) { throw ('Native ' + $Mode + ' failed; exit code ' + $LASTEXITCODE + '. No package was published.') }
}
function Show-BuildMenu($Options) {
    if ([Console]::IsInputRedirected -or [Console]::IsOutputRedirected) {
        throw 'Interactive input requires a console. Use: build.bat package -Theme win98'
    }
    while ($true) {
        Write-Host ('Format: {0} / Data: {1} / Core tests: {2}' -f $Options.Format, $Options.Data, $Options.Tests)
        Write-Host 'F: format  D: sample data  T: core tests  Enter: build  Esc/Q: cancel'
        if ($SkipValidation) { Write-Warning 'Native compile verification will be explicitly skipped.' }
        switch ([Console]::ReadKey($true).Key.ToString()) {
            'F' {
                $formats = @('both','folder','zip')
                $Options.Format = $formats[([Array]::IndexOf($formats, $Options.Format) + 1) % $formats.Count]
            }
            'D' { if ($Options.Data -eq 'sample') { $Options.Data = 'none' } else { $Options.Data = 'sample' } }
            'T' { $Options.Tests = -not $Options.Tests }
            'Escape' { return $null }
            'Q' { return $null }
            'Enter' {
                if ($SkipValidation -and $Options.Tests) { Write-Warning 'Core tests cannot be combined with -SkipValidation.' }
                else { return $Options }
            }
        }
    }
}
function Copy-SafeFile([string]$Source, [string]$Destination) {
    $file = Get-Item -LiteralPath $Source -ErrorAction Stop
    if ($file.PSIsContainer -or ($file.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw ('Invalid source file: ' + $Source) }
    $parent = $file.Directory
    while ($null -ne $parent) {
        if (($parent.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw ('Refusing linked input directory: ' + $parent.FullName) }
        if ($parent.FullName.Equals($script:Root, [StringComparison]::OrdinalIgnoreCase)) { break }
        $parent = $parent.Parent
    }
    Copy-Item -LiteralPath $Source -Destination $Destination
}
function Get-PackageHashes([string]$Directory) {
    $entries = @()
    foreach ($file in (Get-ChildItem -LiteralPath $Directory -Recurse -File | Sort-Object FullName)) {
        $relative = $file.FullName.Substring($Directory.Length + 1).Replace('\','/')
        $entries += [ordered]@{path = $relative; sha256 = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash.ToLowerInvariant()}
    }
    return $entries
}
function Remove-OwnedDirectory([string]$Path, [string]$Parent) {
    $resolved = [IO.Path]::GetFullPath($Path)
    $owner = [IO.Path]::GetFullPath($Parent).TrimEnd([IO.Path]::DirectorySeparatorChar)
    if (-not $resolved.StartsWith($owner + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
        throw ('Cleanup escaped the owned build directory: ' + $resolved)
    }
    if (Test-Path -LiteralPath $resolved) { Remove-Item -LiteralPath $resolved -Recurse -Force }
}
function New-ProductPackage($Options) {
    $compileStatus = 'not_run_explicit_skip'; $testStatus = 'not_requested'
    if ($SkipValidation -and $Options.Tests) { throw '-SkipValidation and -RunTests are mutually exclusive.' }
    if (-not $SkipValidation) {
        Invoke-NativeCheck 'compile'; $compileStatus = 'passed'
        if ($Options.Tests) { Invoke-NativeCheck 'test'; $testStatus = 'passed' }
    } else { Write-Host 'WARNING: Native Windows verification SKIPPED. These packages are NOT runtime-verified.' -ForegroundColor Yellow }
    $destination = $OutputRoot
    if ([string]::IsNullOrWhiteSpace($destination)) { $destination = Join-Path $script:Root 'build/packages' }
    $destination = [IO.Path]::GetFullPath($destination)
    $volumeRoot = [IO.Path]::GetPathRoot($destination)
    if ($destination.Length -gt $volumeRoot.Length) {
        $destination = $destination.TrimEnd([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar)
    }
    # Never write generated output into an input tree. No existing build is removed.
    foreach ($name in @('', 'src', 'web', 'lib', 'data', 'tests', 'tools', 'design', 'docs', 'manual', 'configs', 'samples', 'archive')) {
        $protected = $script:Root
        if ($name) { $protected = Join-Path $protected $name }
        if ($destination.Equals($protected, [StringComparison]::OrdinalIgnoreCase) -or
            ($name -and $destination.StartsWith($protected + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase))) {
            throw ('OutputRoot must not be an application input directory: ' + $destination)
        }
    }
    [IO.Directory]::CreateDirectory($destination) | Out-Null
    $key = [Guid]::NewGuid().ToString('N')
    $stage = Join-Path $destination ('.rdv-stage-' + $key)
    $published = Join-Path $destination ('build-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '-' + $key.Substring(0,8))
    [IO.Directory]::CreateDirectory($stage) | Out-Null
    try {
        $spec = Get-Content -LiteralPath (Join-Path $script:Root 'tools/package-files.json') -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($spec.schema -ne 1 -or $spec.variant -notin @('json-layout','fixed-layout') -or
            $spec.packageName -ne ('ReaderDataViewer-' + $spec.variant)) { throw 'Invalid package product definition.' }
        $id = 'win98'
        $packageName = $spec.packageName
        $package = Join-Path $stage $packageName
        [IO.Directory]::CreateDirectory($package) | Out-Null
        $seen = @{}
        foreach ($file in $spec.files) {
            if ($file -notmatch '^[A-Za-z0-9_./-]+$' -or $file -match '(^/|(^|/)\.\.(/|$))' -or $seen.ContainsKey($file)) {
                throw ('Invalid or duplicate package path: ' + $file)
            }
            $seen[$file] = $true
            $target = Join-Path $package $file
            [IO.Directory]::CreateDirectory((Split-Path -Parent $target)) | Out-Null
            Copy-SafeFile (Join-Path $script:Root $file) $target
        }
        $configSource = Join-Path $script:Root 'configs/sample/settings.json'
        $sampleDirectory = Join-Path $script:Root 'samples/current/data'
        $sampleSet = 'generic-five'
        if ($SampleRoot) {
            $sourceRoot = (Get-Item -LiteralPath $SampleRoot -ErrorAction Stop).FullName
            $configName = 'settings.json'
            if ($spec.variant -eq 'fixed-layout') { $configName = 'settings-fixed.json' }
            $configSource = Join-Path $sourceRoot $configName
            $sampleDirectory = Join-Path $sourceRoot 'サンプル/data'
            $sampleSet = 'selected-five'
        }
        $config = Get-Content -LiteralPath $configSource -Raw -Encoding UTF8 | ConvertFrom-Json
        $sampleSourceNames = @{}
        $sampleKeys = @('TXN','PAY','APP','DEL')
        for ($sampleIndex = 0; $sampleIndex -lt $sampleKeys.Count; $sampleIndex++) {
            $table = $config.data.tables.($sampleKeys[$sampleIndex])
            $oldName = [string]$table.file
            $newName = $oldName
            if ($oldName -notmatch '^0[1-4]_') { $newName = ('{0:D2}_' -f ($sampleIndex + 1)) + $oldName.TrimStart([char[]]'①②③④') }
            $sampleSourceNames[$newName] = $oldName
            $table.file = $newName
        }
        if ($SampleRoot) {
            $config.paths.dataDir = 'data'
            $config.paths.ledger = 'data/統合台帳.xlsx'
            $config.paths.log = 'data/操作ログ.log'
            $config.watch.targets = @()
            if ($spec.variant -eq 'json-layout') {
                $statusBar = @($config.screen.sections | Where-Object { $_.type -eq 'statusBar' })[0]
                if (-not @($statusBar.buttons | Where-Object { $_.action -eq 'restoreRecords' }).Count) {
                    $statusBar.buttons = @($statusBar.buttons | Where-Object { $_.action -ne 'settings' }) + @([pscustomobject]@{action='restoreRecords';text='削除済み'}) + @($statusBar.buttons | Where-Object { $_.action -eq 'settings' })
                }
            }
            Write-Utf8 (Join-Path $package 'settings.json') (($config | ConvertTo-Json -Depth 50) + [Environment]::NewLine)
        } else {
            Copy-SafeFile $configSource (Join-Path $package 'settings.json')
        }
        [IO.Directory]::CreateDirectory((Join-Path $package 'data')) | Out-Null
        if ($Options.Data -eq 'sample') {
            $names = @($config.data.tables.PSObject.Properties | ForEach-Object { [string]$_.Value.file })
            if ($names.Count -ne 4 -or @($names | Sort-Object -Unique).Count -ne 4) { throw 'The five-record sample must have four distinct inputs.' }
            foreach ($name in $names) {
                if ([IO.Path]::GetFileName($name) -ne $name -or $name -match '[/\\:]' -or [IO.Path]::GetExtension($name) -notin @('.csv','.xlsx')) { throw ('Invalid sample input filename: ' + $name) }
                Copy-SafeFile (Join-Path $sampleDirectory $sampleSourceNames[$name]) (Join-Path $package ('data/' + $name))
            }
        }
        if ($Options.Data -eq 'sample' -and -not $SampleRoot) {
            $updateName = '追加CSVデータ/5月分（3件追加）'
            $updateSource = Join-Path (Join-Path $script:Root 'samples/current') $updateName
            $updateTarget = Join-Path $package $updateName
            [IO.Directory]::CreateDirectory($updateTarget) | Out-Null
            $updateConfig = Get-Content -LiteralPath (Join-Path $script:Root 'configs/sample/settings.json') -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($table in $updateConfig.data.tables.PSObject.Properties) {
                $name = [string]$table.Value.file
                if ([IO.Path]::GetFileName($name) -ne $name) { throw 'Invalid update input filename' }
                Copy-SafeFile (Join-Path $updateSource $name) (Join-Path $updateTarget $name)
            }
        }
        $sourceCommit = $null; $sourceDirty = $null
        if ((Test-Path -LiteralPath (Join-Path $script:Root '.git')) -and (Get-Command git -ErrorAction SilentlyContinue)) {
            $revision = & git -C $script:Root rev-parse HEAD 2>$null
            if ($LASTEXITCODE -eq 0) {
                $sourceCommit = [string]$revision
                $status = & git -C $script:Root status --porcelain --untracked-files=normal 2>$null
                if ($LASTEXITCODE -eq 0) { $sourceDirty = [bool]$status }
            }
        }
        $release = [ordered]@{
            schema=1; application='ReaderDataViewer'; variant=$spec.variant; name=$spec.name
            createdUtc=[DateTime]::UtcNow.ToString('o'); sourceCommit=$sourceCommit; sourceDirty=$sourceDirty; data=$Options.Data; sampleSet=$sampleSet
        }
        Write-Json (Join-Path $package 'src/release.json') $release
        $manifest = [ordered]@{
            schema = 1; application = 'ReaderDataViewer'; variant = $spec.variant; name = $spec.name; sourceCommit = $sourceCommit; sourceDirty = $sourceDirty; theme = $id; motion = 'off'
            createdUtc = [DateTime]::UtcNow.ToString('o'); data = $Options.Data
            validation = [ordered]@{ nativeCompile = $compileStatus; coreTests = $testStatus; windowsUI = 'not_run_by_packager'; sharedLedger = 'not_run_by_packager' }
            files = @(Get-PackageHashes $package)
        }
        if ($Options.Format -ne 'folder') {
            Add-Type -AssemblyName System.IO.Compression,System.IO.Compression.FileSystem
            $zipPath = Join-Path $stage ($packageName + '.zip')
            $zip = [IO.Compression.ZipFile]::Open($zipPath, [IO.Compression.ZipArchiveMode]::Create)
            try {
                foreach ($directory in (Get-ChildItem -LiteralPath $package -Recurse -Directory)) {
                    if (@(Get-ChildItem -LiteralPath $directory.FullName -Force).Count -eq 0) {
                        $zip.CreateEntry($packageName + '/' + $directory.FullName.Substring($package.Length + 1).Replace('\','/') + '/') | Out-Null
                    }
                }
                foreach ($file in (Get-ChildItem -LiteralPath $package -Recurse -File | Sort-Object FullName)) {
                    $entryName = $packageName + '/' + $file.FullName.Substring($package.Length + 1).Replace('\','/')
                    [IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $file.FullName, $entryName, [IO.Compression.CompressionLevel]::Optimal) | Out-Null
                }
            } finally { $zip.Dispose() }
        }
        if ($Options.Format -eq 'zip') { Remove-OwnedDirectory $package $stage }
        $summary = @([ordered]@{theme=$id; variant=$spec.variant; name=$spec.name; motion='off'; package=$packageName; format=$Options.Format})
        Write-Host ('  Prepared: ' + $packageName) -ForegroundColor Green
        Write-Json (Join-Path $stage 'package-manifest.json') $manifest
        Write-Json (Join-Path $stage 'build-summary.json') ([ordered]@{schema=1; nativeCompile=$compileStatus; coreTests=$testStatus; packages=$summary})
        [IO.Directory]::Move($stage, $published)
        Write-Host ''
        Write-Host ('Published: ' + $published) -ForegroundColor Cyan
        if ($compileStatus -ne 'passed') { Write-Host 'Windows verification was NOT performed. See package-manifest.json.' -ForegroundColor Yellow }
    }
    finally {
        # This path is an owned GUID staging directory, never a supplied destination.
        Remove-OwnedDirectory $stage $destination
    }
}
try {
    if ($Command -eq 'help') { Show-Usage; exit 0 }
    if ($Command -eq 'compile' -or $Command -eq 'test') {
        if ($Theme -or $All -or $SkipValidation -or $RunTests -or $OutputRoot -or $SampleRoot -or
            $PSBoundParameters.ContainsKey('Format') -or $PSBoundParameters.ContainsKey('Data')) {
            throw 'compile/test do not accept packaging options.'
        }
        Invoke-NativeCheck $Command; exit 0
    }
    if ($All -and $Theme) { throw '-All and -Theme are mutually exclusive.' }
    if ($Theme -and $Theme.Trim() -ne 'win98') { throw 'Win98 is the only supported theme.' }
    if ($Command -eq 'list') {
        $spec = Get-Content -LiteralPath (Join-Path $script:Root 'tools/package-files.json') -Raw -Encoding UTF8 | ConvertFrom-Json
        [pscustomobject]@{id=$spec.variant; name=$spec.name; package=$spec.packageName} | Format-Table -AutoSize
        exit 0
    }
    $options = [pscustomobject]@{Format = $Format; Data = $Data; Tests = [bool]$RunTests}
    if ($Command -eq 'menu') {
        $options = Show-BuildMenu $options
        if ($null -eq $options) { Write-Host 'Cancelled. No files were written.'; exit 0 }
    }
    New-ProductPackage $options
    exit 0
}
catch {
    [Console]::Error.WriteLine('BUILD FAILED: ' + $_.Exception.Message)
    exit 1
}
