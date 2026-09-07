# Windows PowerShell 5.1 compatible; no npm, SDK, downloaded UI kit or admin needed.
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [ValidateSet('menu', 'compile', 'test', 'package', 'list', 'help')]
    [string]$Command = 'menu',
    [Alias('Themes')][string]$Theme = '',
    [switch]$All,
    [ValidateSet('auto', 'off')][string]$Motion = 'auto',
    [ValidateSet('both', 'folder', 'zip')][string]$Format = 'both',
    [ValidateSet('sample', 'none')][string]$Data = 'sample',
    [string]$OutputRoot = '',
    [switch]$RunTests,
    [switch]$SkipValidation
)
Set-StrictMode -Version 2
$ErrorActionPreference = 'Stop'
$script:Root = Split-Path -Parent $PSScriptRoot
$script:Utf8 = New-Object System.Text.UTF8Encoding($false)

function Write-Utf8([string]$Path, [string]$Text) {
    [IO.File]::WriteAllText($Path, $Text, $script:Utf8)
}
function Write-Json([string]$Path, $Object) {
    Write-Utf8 $Path (($Object | ConvertTo-Json -Depth 10) + [Environment]::NewLine)
}
function Show-Usage {
    @'
Reader Data Viewer - design build console

  build.bat                         Interactive multi-select TUI
  build.bat compile                 Original C# compile-only check
  build.bat test                    Original isolated C# regression tests
  build.bat list                    List the six designs
  build.bat package -All            Validate and package all six designs
  build.bat package -Theme apple,fluent -Motion off
  build.bat package -Theme win98 -Format zip -RunTests
  build.bat package -All -Data none -OutputRoot "C:\RDV releases"

Options:
  -Theme ID[,ID...]  win98 / apple / material / fluent / carbon / spectrum
  -All              Select all six (do not combine with -Theme)
  -Motion auto|off  Short non-blocking feedback; classic is always off
  -Format both|folder|zip   Output format (default: both)
  -Data sample|none  Bundled test CSVs only, NEVER live ledger/log/pending data
  -RunTests         Run isolated core regression tests before packaging
  -SkipValidation   Explicitly skip Windows compile verification (recorded)
                    Cannot be combined with -RunTests; not a production pass.

Each build creates a NEW timestamped folder under releases (or -OutputRoot).
Existing packages and input files are never overwritten.
This is the original source-at-startup application, not a standalone .exe.
See docs/design-build.md for Japanese instructions and deployment precautions.
'@ | Write-Host
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
function Read-ThemeIds([string]$Text, $Catalog) {
    $known = @($Catalog | ForEach-Object { $_.id })
    $ids = @($Text.ToLowerInvariant().Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Select-Object -Unique)
    if ($ids.Count -eq 0) { throw 'Select at least one theme with -Theme or -All.' }
    foreach ($id in $ids) { if ($known -notcontains $id) { throw ('Unknown theme: ' + $id) } }
    return $ids
}
function Read-LineMenu($Catalog, [string]$Initial) {
    Write-Host 'Line-input mode. Choose one or more theme numbers (comma separated).'
    for ($i = 0; $i -lt $Catalog.Count; $i++) { Write-Host ('  {0}. {1}' -f ($i + 1), $Catalog[$i].name) }
    while ($true) {
        $text = Read-Host ('Numbers / A=all / Q=cancel [default: ' + $Initial + ']')
        if ($text -match '^[qQ]$') { return $null }
        if ([string]::IsNullOrWhiteSpace($text)) { return @($Initial.Split(',')) }
        if ($text -match '^[aA]$') { return @($Catalog | ForEach-Object { $_.id }) }
        $ids = @(); $valid = $true
        foreach ($part in $text.Split(',')) {
            $index = 0
            if (-not [int]::TryParse($part.Trim(), [ref]$index) -or $index -lt 1 -or $index -gt $Catalog.Count) { $valid = $false; break }
            $ids += $Catalog[$index - 1].id
        }
        if ($valid -and $ids.Count -gt 0) { return @($ids | Select-Object -Unique) }
        Write-Host 'Enter numbers 1-6 separated by commas.' -ForegroundColor Yellow
    }
}
function Show-BuildMenu($Catalog, [string[]]$Initial) {
    if ([Console]::IsInputRedirected -or [Console]::IsOutputRedirected) {
        throw 'Interactive input requires a console. Use: build.bat package -All (or -Theme ID,ID).'
    }
    $chosen = @{}
    foreach ($id in $Initial) { $chosen[$id] = $true }
    $cursor = 0; $message = ''
    $options = [ordered]@{ Ids = @(); Motion = $Motion; Format = $Format; Data = $Data; Tests = [bool]$RunTests }
    $oldColour = [Console]::ForegroundColor
    $oldCursor = [Console]::CursorVisible
    try {
        # Remote hosts without raw-key/clear support get a non-animated line menu.
        if ([Console]::WindowWidth -lt 64 -or [Console]::WindowHeight -lt 22) {
            $ids = Read-LineMenu $Catalog ($Initial -join ',')
            if ($null -eq $ids) { return $null }
            $options.Ids = @($ids); return [pscustomobject]$options
        }
        [Console]::CursorVisible = $false
        while ($true) {
            try { [Console]::Clear() } catch {
                [Console]::CursorVisible = $oldCursor
                $ids = Read-LineMenu $Catalog ($Initial -join ',')
                if ($null -eq $ids) { return $null }
                $options.Ids = @($ids); return [pscustomobject]$options
            }
            Write-Host ' READER DATA VIEWER  /  DESIGN BUILD' -ForegroundColor Cyan
            Write-Host ' --------------------------------------------------------------'
            Write-Host ' Up/Down: move  Space/1-6: select  A: all/clear  Enter: build'
            Write-Host ''
            for ($i = 0; $i -lt $Catalog.Count; $i++) {
                $mark = ' '; if ($chosen.ContainsKey($Catalog[$i].id)) { $mark = 'x' }
                $pointer = ' '; if ($i -eq $cursor) { $pointer = '>' }
                $colour = 'Gray'; if ($i -eq $cursor) { $colour = 'Cyan' }
                Write-Host (' {0} [{1}] {2}. {3}' -f $pointer, $mark, ($i + 1), $Catalog[$i].name) -ForegroundColor $colour
            }
            Write-Host ''
            Write-Host (' ' + $Catalog[$cursor].description)
            Write-Host (' M: Motion = {0}  (Win98 always OFF)' -f $options.Motion)
            Write-Host (' F: Format = {0}   D: Data = {1}' -f $options.Format, $options.Data)
            Write-Host (' T: Core tests = {0}  | C# compile check: required' -f $options.Tests)
            if ($SkipValidation) { Write-Host ' WARNING: -SkipValidation was supplied; compile is NOT checked.' -ForegroundColor Yellow }
            Write-Host ' Esc / Q: cancel without writing files'
            Write-Host ' Output: a NEW build folder. No ledger/log/pending files copied.'
            if ($message) { Write-Host (' ' + $message) -ForegroundColor Yellow }
            $key = [Console]::ReadKey($true)
            $toggle = -1
            switch ($key.Key.ToString()) {
                'UpArrow' { $cursor = ($cursor + $Catalog.Count - 1) % $Catalog.Count }
                'DownArrow' { $cursor = ($cursor + 1) % $Catalog.Count }
                'Spacebar' { $toggle = $cursor }
                'A' {
                    if ($chosen.Count -eq $Catalog.Count) { $chosen.Clear() }
                    else { foreach ($item in $Catalog) { $chosen[$item.id] = $true } }
                }
                'M' { if ($options.Motion -eq 'auto') { $options.Motion = 'off' } else { $options.Motion = 'auto' } }
                'F' { $formats = @('both','folder','zip'); $options.Format = $formats[([Array]::IndexOf($formats, $options.Format) + 1) % 3] }
                'D' { if ($options.Data -eq 'sample') { $options.Data = 'none' } else { $options.Data = 'sample' } }
                'T' { $options.Tests = -not $options.Tests }
                'Escape' { return $null }
                'Q' { return $null }
                'Enter' {
                    if ($chosen.Count -eq 0) { $message = 'Select at least one theme.' }
                    elseif ($SkipValidation -and $options.Tests) { $message = 'Core tests cannot be combined with -SkipValidation.' }
                    else {
                        $options.Ids = @($Catalog | Where-Object { $chosen.ContainsKey($_.id) } | ForEach-Object { $_.id })
                        return [pscustomobject]$options
                    }
                }
            }
            if ([string]$key.KeyChar -match '^[1-6]$') { $toggle = [int]::Parse([string]$key.KeyChar) - 1 }
            if ($toggle -ge 0 -and $toggle -lt $Catalog.Count) {
                $id = $Catalog[$toggle].id
                if ($chosen.ContainsKey($id)) { $chosen.Remove($id) } else { $chosen[$id] = $true }
                $cursor = $toggle; $message = ''
            }
        }
    }
    finally { [Console]::ForegroundColor = $oldColour; [Console]::CursorVisible = $oldCursor }
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
function New-DesignPackages($Catalog, $Options) {
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
        $summary = @()
        foreach ($id in $Options.Ids) {
            $definition = @($Catalog | Where-Object { $_.id -eq $id })[0]
            $packageName = 'ReaderDataViewer-' + $id
            $package = Join-Path $stage $packageName
            [IO.Directory]::CreateDirectory($package) | Out-Null
            foreach ($file in @('ReaderDataViewer.cmd', 'ReaderDataViewer.vbs', 'settings.json', 'LICENSE', 'THIRD-PARTY-NOTICES.md')) {
                Copy-SafeFile (Join-Path $script:Root $file) (Join-Path $package $file)
            }
            foreach ($dir in @('src','web','lib')) { Copy-SafeTree (Join-Path $script:Root $dir) (Join-Path $package $dir) }
            [IO.Directory]::CreateDirectory((Join-Path $package 'data')) | Out-Null
            [IO.Directory]::CreateDirectory((Join-Path $package 'output')) | Out-Null
            if ($Options.Data -eq 'sample') {
                # Deliberate allow-list: NEVER copy the active data/ directory.
                foreach ($csv in @('tableA.csv','tableB.csv','tableC.csv','delete.csv')) {
                    Copy-SafeFile (Join-Path $script:Root ('tests/fixtures/data/' + $csv)) (Join-Path $package ('data/' + $csv))
                }
            }
            $effectiveMotion = 'off'; if ($definition.modern) { $effectiveMotion = $Options.Motion }
            $config = [ordered]@{schema = 1; id = $id; motion = $effectiveMotion; background = $definition.background; caption = $definition.caption; captionText = $definition.captionText; border = $definition.border}
            Write-Json (Join-Path $package 'web/theme.json') $config
            $htmlPath = Join-Path $package 'web/index.html'
            $html = [IO.File]::ReadAllText($htmlPath, [Text.Encoding]::UTF8)
            $attributes = [ordered]@{'data-rdv-theme' = $id; 'data-rdv-modern' = ([string][bool]$definition.modern).ToLowerInvariant(); 'data-rdv-motion' = $effectiveMotion}
            foreach ($attribute in $attributes.Keys) {
                $pattern = '\b' + $attribute + '="[^"]*"'
                if ([regex]::Matches($html, $pattern).Count -ne 1) { throw ('Expected exactly one theme marker: ' + $attribute) }
                $html = [regex]::Replace($html, $pattern, ($attribute + '="' + $attributes[$attribute] + '"'))
            }
            Write-Utf8 $htmlPath $html
            [IO.Directory]::CreateDirectory((Join-Path $package 'docs')) | Out-Null
            foreach ($doc in @('design-build.md','theme-validation.md','settings.md','shared-ledger.md')) {
                Copy-SafeFile (Join-Path $script:Root ('docs/' + $doc)) (Join-Path $package ('docs/' + $doc))
            }
            $readme = "Reader Data Viewer - $($definition.name)`r`n`r`n" +
                "Start: ReaderDataViewer.vbs (or .cmd for console diagnostics).`r`n" +
                "Extract the entire ZIP first. Requires 64-bit Windows, Windows PowerShell 5.1, WPF and WebView2 Runtime.`r`n" +
                "Theme: $id / motion: $effectiveMotion / input data: $($Options.Data)`r`n" +
                "Native compile: $compileStatus / core tests: $testStatus`r`n" +
                "This is a source-at-startup distribution, NOT a standalone EXE.`r`n" +
                "The original business settings are copied unchanged. Sample data is not your live data.`r`n" +
                "No live ledger, log, output or local pending changes were copied.`r`n" +
                "Review paths in settings.json BEFORE running a production copy.`r`n" +
                "Use one deployed theme per working folder. Close the app and back up data before switching themes.`r`n" +
                "Japanese instructions: docs/design-build.md (source build tools are in the developer ZIP).`r`n" +
                "Vendor names describe design inspiration, not certification or vendor component binaries.`r`n"
            Write-Utf8 (Join-Path $package 'PACKAGE-README.txt') $readme
            $manifest = [ordered]@{
                schema = 1; application = 'ReaderDataViewer'; theme = $id; motion = $effectiveMotion
                createdUtc = [DateTime]::UtcNow.ToString('o'); data = $Options.Data
                validation = [ordered]@{ nativeCompile = $compileStatus; coreTests = $testStatus; windowsUI = 'not_run_by_packager'; sharedLedger = 'not_run_by_packager' }
                files = @(Get-PackageHashes $package)
            }
            Write-Json (Join-Path $package 'package-manifest.json') $manifest
            if ($Options.Format -ne 'folder') {
                Add-Type -AssemblyName System.IO.Compression.FileSystem
                [IO.Compression.ZipFile]::CreateFromDirectory($package, (Join-Path $stage ($packageName + '.zip')), [IO.Compression.CompressionLevel]::Optimal, $true)
            }
            if ($Options.Format -eq 'zip') { Remove-Item -LiteralPath $package -Recurse -Force }
            $summary += [ordered]@{theme=$id; name=$definition.name; motion=$effectiveMotion; package=$packageName; format=$Options.Format}
            Write-Host ('  Prepared: ' + $definition.name) -ForegroundColor Green
        }
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
            $PSBoundParameters.ContainsKey('Motion') -or $PSBoundParameters.ContainsKey('Format') -or $PSBoundParameters.ContainsKey('Data')) {
            throw 'compile/test do not accept packaging options.'
        }
        Invoke-NativeCheck $Command; exit 0
    }
    $document = [IO.File]::ReadAllText((Join-Path $script:Root 'design/themes.json'), [Text.Encoding]::UTF8) | ConvertFrom-Json
    $catalog = @($document.themes)
    if ($document.schema -ne 1 -or $catalog.Count -ne 6 -or @($catalog.id | Select-Object -Unique).Count -ne 6) { throw 'Invalid theme catalog.' }
    $expectedIds = @('win98','apple','material','fluent','carbon','spectrum')
    foreach ($entry in $catalog) {
        if ($expectedIds -notcontains $entry.id) { throw 'Invalid theme id.' }
        if ($entry.modern -isnot [bool] -or $entry.modern -ne ($entry.id -ne 'win98')) { throw 'Invalid theme mode.' }
        foreach ($colour in @($entry.background,$entry.caption,$entry.captionText,$entry.border)) { if ($colour -notmatch '^#[a-fA-F0-9]{6}$') { throw 'Invalid theme colour.' } }
    }
    if ($Command -eq 'list') { $catalog | Format-Table id, name, modern -AutoSize; exit 0 }
    if ($All -and $Theme) { throw '-All and -Theme are mutually exclusive.' }
    $ids = @('win98')
    if ($All) { $ids = @($catalog | ForEach-Object { $_.id }) }
    elseif ($Theme) { $ids = @(Read-ThemeIds $Theme $catalog) }
    elseif ($Command -eq 'package') { throw 'package requires -All or -Theme ID[,ID...].' }
    $options = [pscustomobject]@{Ids = $ids; Motion = $Motion; Format = $Format; Data = $Data; Tests = [bool]$RunTests}
    if ($Command -eq 'menu') {
        $options = Show-BuildMenu $catalog $ids
        if ($null -eq $options) { Write-Host 'Cancelled. No files were written.'; exit 0 }
    }
    New-DesignPackages $catalog $options
    exit 0
}
catch {
    [Console]::Error.WriteLine('BUILD FAILED: ' + $_.Exception.Message)
    exit 1
}
