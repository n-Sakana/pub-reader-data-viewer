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
        'Reader Data Viewer - Win98 build'
        ''
        '  build.bat                         Choose format, sample data and tests'
        '  build.bat compile                 Compile the C# application'
        '  build.bat test                    Run isolated C# regression tests'
        '  build.bat package -Theme win98     Create a new folder and ZIP'
        '  build.bat package -Format zip -Data none -OutputRoot "C:\RDV releases"'
        ''
        'Options: -Format both|folder|zip, -Data sample|none, -OutputRoot PATH'
        '         -RunTests, -SkipValidation (explicit compile skip, never a pass)'
        'Win98 is the only design. -Theme win98 and -All are optional aliases.'
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
        Write-Host ('Win98 / Format: {0} / Data: {1} / Core tests: {2}' -f $Options.Format, $Options.Data, $Options.Tests)
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
function Copy-SafeTree([string]$Source, [string]$Destination) {
    if (-not (Test-Path -LiteralPath $Source -PathType Container)) { throw ('Missing directory: ' + $Source) }
    $items = @(Get-Item -LiteralPath $Source) + @(Get-ChildItem -LiteralPath $Source -Force -Recurse)
    foreach ($item in $items) {
        if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw ('Refusing linked/reparse-point input: ' + $item.FullName) }
    }
    [IO.Directory]::CreateDirectory($Destination) | Out-Null
    foreach ($item in (Get-ChildItem -LiteralPath $Source -Force)) {
        Copy-Item -LiteralPath $item.FullName -Destination $Destination -Recurse -Force
    }
}
function Copy-SafeFile([string]$Source, [string]$Destination) {
    $file = Get-Item -LiteralPath $Source -ErrorAction Stop
    if ($file.PSIsContainer -or ($file.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw ('Invalid source file: ' + $Source) }
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
function New-Win98Package($Options) {
    $compileStatus = 'not_run_explicit_skip'; $testStatus = 'not_requested'
    if ($SkipValidation -and $Options.Tests) { throw '-SkipValidation and -RunTests are mutually exclusive.' }
    if (-not $SkipValidation) {
        Invoke-NativeCheck 'compile'; $compileStatus = 'passed'
        if ($Options.Tests) { Invoke-NativeCheck 'test'; $testStatus = 'passed' }
    } else { Write-Host 'WARNING: Native Windows verification SKIPPED. These packages are NOT runtime-verified.' -ForegroundColor Yellow }
    $destination = $OutputRoot
    if ([string]::IsNullOrWhiteSpace($destination)) { $destination = Join-Path $script:Root 'releases' }
    $destination = [IO.Path]::GetFullPath($destination)
    $volumeRoot = [IO.Path]::GetPathRoot($destination)
    if ($destination.Length -gt $volumeRoot.Length) {
        $destination = $destination.TrimEnd([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar)
    }
    # Never write generated output into an input tree. No existing build is removed.
    foreach ($name in @('', 'src', 'web', 'lib', 'data', 'tests', 'tools', 'design', 'docs')) {
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
        $id = 'win98'
        $packageName = 'ReaderDataViewer-' + $id
        $package = Join-Path $stage $packageName
        [IO.Directory]::CreateDirectory($package) | Out-Null
        foreach ($file in @('ReaderDataViewer.cmd', 'ReaderDataViewer.vbs', 'README.md', 'PAYMENT-GUIDE.md', 'LICENSE', 'THIRD-PARTY-NOTICES.md')) {
            Copy-SafeFile (Join-Path $script:Root $file) (Join-Path $package $file)
        }
        # 見本データ一式と同じ形の設定を入れる。sample-v4 に無ければ直下のものを使う。
        $sampleSettings = Join-Path $script:Root 'tests/fixtures/sample-v4/settings.json'
        if ($Options.Data -eq 'sample' -and (Test-Path -LiteralPath $sampleSettings -PathType Leaf)) {
            Copy-SafeFile $sampleSettings (Join-Path $package 'settings.json')
        } else {
            Copy-SafeFile (Join-Path $script:Root 'settings.json') (Join-Path $package 'settings.json')
        }
        foreach ($dir in @('src','web','lib')) { Copy-SafeTree (Join-Path $script:Root $dir) (Join-Path $package $dir) }
        [IO.Directory]::CreateDirectory((Join-Path $package 'data')) | Out-Null
        [IO.Directory]::CreateDirectory((Join-Path $package 'output')) | Out-Null
        if ($Options.Data -eq 'sample') {
            # Deliberate allow-list: NEVER copy the active data/ directory.
            # Sample CSVs are a convenience, not a build requirement. A missing
            # fixture warns and produces a package without sample data; it never
            # fails the build.
            $missing = @()
            $sampleDir = Join-Path $script:Root 'tests/fixtures/sample-v4'
            foreach ($name in @(
                '①取引データ_100件.csv',
                '②決済管理データ_100件.csv',
                '③講習受講データ_100件.xlsx',
                '④処理済みデータ_100件.xlsx')) {
                $fixture = Join-Path $sampleDir $name
                if (Test-Path -LiteralPath $fixture -PathType Leaf) {
                    Copy-SafeFile $fixture (Join-Path $package ('data/' + $name))
                } else {
                    $missing += $name
                }
            }
            if ($missing.Count -gt 0) {
                Write-Warning ('Sample data not found in tests/fixtures/sample-v4: ' + ($missing -join ', ') + '. The package is built without it.')
            }
        }
        # docs/ は開発中の記録なので配布しない (先生の指示 2026-09-10)。
        # 実機名や検証の経緯が入っていて、受け取る人には要らない。
        $readme = "Reader Data Viewer - Windows 98 Classic`r`n`r`n" +
            "Start: ReaderDataViewer.vbs (or .cmd for console diagnostics).`r`n" +
            "Extract the entire ZIP first. Requires 64-bit Windows, Windows PowerShell 5.1, WPF and WebView2 Runtime.`r`n" +
            "Theme: win98 / motion: off / input data: $($Options.Data)`r`n" +
            "Native compile: $compileStatus / core tests: $testStatus`r`n" +
            "This is a source-at-startup distribution, NOT a standalone EXE.`r`n" +
            "The sample uses PAY+MAP pairs, two-status payment checks and processed-only deletion. See PAYMENT-GUIDE.md.`r`n" +
            "No live ledger, log, output or local pending changes were copied.`r`n" +
            "Review paths in settings.json BEFORE running a production copy.`r`n" +
            "Check steps before acceptance: README.md`r`n"
        Write-Utf8 (Join-Path $package 'PACKAGE-README.txt') $readme
        $manifest = [ordered]@{
            schema = 1; application = 'ReaderDataViewer'; theme = $id; motion = 'off'
            createdUtc = [DateTime]::UtcNow.ToString('o'); data = $Options.Data
            validation = [ordered]@{ nativeCompile = $compileStatus; coreTests = $testStatus; windowsUI = 'not_run_by_packager'; sharedLedger = 'not_run_by_packager' }
            files = @(Get-PackageHashes $package)
        }
        if ($Options.Format -ne 'folder') {
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            [IO.Compression.ZipFile]::CreateFromDirectory($package, (Join-Path $stage ($packageName + '.zip')), [IO.Compression.CompressionLevel]::Optimal, $true)
        }
        if ($Options.Format -eq 'zip') { Remove-Item -LiteralPath $package -Recurse -Force }
        $summary = @([ordered]@{theme=$id; name='Windows 98 Classic'; motion='off'; package=$packageName; format=$Options.Format})
        Write-Host '  Prepared: Windows 98 Classic' -ForegroundColor Green
        Write-Json (Join-Path $stage 'package-manifest.json') $manifest
        Write-Json (Join-Path $stage 'package-manifest.json') $manifest
        Write-Json (Join-Path $stage 'build-summary.json') ([ordered]@{schema=1; nativeCompile=$compileStatus; coreTests=$testStatus; packages=$summary})
        [IO.Directory]::Move($stage, $published)
        Write-Host ''
        Write-Host ('Published: ' + $published) -ForegroundColor Cyan
        if ($compileStatus -ne 'passed') { Write-Host 'Windows verification was NOT performed. See package-manifest.json.' -ForegroundColor Yellow }
    }
    finally {
        # This path is an owned GUID staging directory, never a supplied destination.
        if (Test-Path -LiteralPath $stage) { Remove-Item -LiteralPath $stage -Recurse -Force }
    }
}
try {
    if ($Command -eq 'help') { Show-Usage; exit 0 }
    if ($Command -eq 'compile' -or $Command -eq 'test') {
        if ($Theme -or $All -or $SkipValidation -or $RunTests -or $OutputRoot -or
            $PSBoundParameters.ContainsKey('Format') -or $PSBoundParameters.ContainsKey('Data')) {
            throw 'compile/test do not accept packaging options.'
        }
        Invoke-NativeCheck $Command; exit 0
    }
    if ($All -and $Theme) { throw '-All and -Theme are mutually exclusive.' }
    if ($Theme -and $Theme.Trim() -ne 'win98') { throw 'Win98 is the only supported theme.' }
    if ($Command -eq 'list') {
        [pscustomobject]@{id='win98'; name='Windows 98 Classic'; modern=$false} | Format-Table -AutoSize
        exit 0
    }
    $options = [pscustomobject]@{Format = $Format; Data = $Data; Tests = [bool]$RunTests}
    if ($Command -eq 'menu') {
        $options = Show-BuildMenu $options
        if ($null -eq $options) { Write-Host 'Cancelled. No files were written.'; exit 0 }
    }
    New-Win98Package $options
    exit 0
}
catch {
    [Console]::Error.WriteLine('BUILD FAILED: ' + $_.Exception.Message)
    exit 1
}
