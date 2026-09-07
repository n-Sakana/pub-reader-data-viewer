[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

# The product is compiled the way ReaderDataViewer.ps1 compiles it: every
# src\*.cs in name order, usings hoisted, one Add-Type under the in-box csc.
# The three steps are separate functions so compile_check.ps1 can concatenate
# the sources itself (it keeps a line map) without a second copy of the
# assembly list.

function Import-RdvWebView2 {
    param([Parameter(Mandatory = $true)][string]$Root)

    $libraryDirectory = Join-Path $Root 'lib'

    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName WindowsBase
    Add-Type -AssemblyName System.Xaml
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName UIAutomationClient
    Add-Type -AssemblyName UIAutomationTypes
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.Xml

    $env:Path = $libraryDirectory + [IO.Path]::PathSeparator + $env:Path
    $webViewAssemblies = @(
        (Join-Path $libraryDirectory 'Microsoft.Web.WebView2.Core.dll')
        (Join-Path $libraryDirectory 'Microsoft.Web.WebView2.Wpf.dll')
    )
    foreach ($assemblyPath in $webViewAssemblies) {
        if (-not (Test-Path -LiteralPath $assemblyPath -PathType Leaf)) {
            throw "Required WebView2 assembly is missing: $assemblyPath"
        }

        # Loaded as bytes, not by path: a zip-borne Mark of the Web would
        # otherwise block the managed DLL.
        [Reflection.Assembly]::Load(
            [IO.File]::ReadAllBytes($assemblyPath)) | Out-Null
    }

    return $webViewAssemblies
}

function Get-RdvReferences {
    param([Parameter(Mandatory = $true)][string[]]$WebViewAssemblies)

    return @(
        [System.Windows.Window].Assembly.Location
        [System.Windows.UIElement].Assembly.Location
        [System.Windows.DependencyObject].Assembly.Location
        [System.Xaml.XamlReader].Assembly.Location
        [System.Windows.Automation.AutomationElement].Assembly.Location
        [System.Windows.Automation.ControlType].Assembly.Location
        [System.IO.Compression.ZipArchive].Assembly.Location
        [System.Xml.XmlDocument].Assembly.Location
        'System.Drawing'
        $WebViewAssemblies[0]
        $WebViewAssemblies[1]
    )
}

function Get-RdvSourceFiles {
    param([Parameter(Mandatory = $true)][string]$Root)

    $sourceDirectory = Join-Path $Root 'src'
    $sourceFiles = @(Get-ChildItem -LiteralPath $sourceDirectory -Filter '*.cs' -File |
        Sort-Object -Property Name)
    if ($sourceFiles.Count -eq 0) {
        throw "No C# source files were found in: $sourceDirectory"
    }

    return $sourceFiles
}

function Import-RdvProduct {
    param([Parameter(Mandatory = $true)][string]$Root)

    $webViewAssemblies = Import-RdvWebView2 -Root $Root
    $sourceFiles = Get-RdvSourceFiles -Root $Root

    $combined = ($sourceFiles | ForEach-Object {
        [IO.File]::ReadAllText($_.FullName, [Text.Encoding]::UTF8)
    }) -join [Environment]::NewLine
    $usingPattern = '(?m)^\s*using\s+[A-Za-z_][A-Za-z0-9_.]*\s*;\s*$'
    $usings = [regex]::Matches($combined, $usingPattern) |
        ForEach-Object { $_.Value.Trim() } |
        Sort-Object -Unique
    $body = [regex]::Replace($combined, $usingPattern, '')
    $source = ($usings -join [Environment]::NewLine) +
        [Environment]::NewLine + [Environment]::NewLine + $body

    Add-Type -TypeDefinition $source `
        -ReferencedAssemblies (Get-RdvReferences -WebViewAssemblies $webViewAssemblies) `
        -Language CSharp

    Write-Output ("compile ok ({0} sources)" -f $sourceFiles.Count)
}

function Remove-RdvTestDirectory {
    <#
      Drop a scratch directory once its check has passed. Each run copies
      src\, lib\, web\ and data\ -- about 2 MB a time -- so keeping every run
      is what grew this tree to gigabytes. A FAILED run keeps its directory:
      that is the evidence.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][bool]$Passed
    )

    if (-not $Passed) {
        Write-Output ("scratch kept for the failure: {0}" -f $Path)
        return
    }
    try {
        Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop
    }
    catch {
        # A file the app still holds is not a test failure; say so and move on.
        Write-Output ("scratch could not be removed: {0} ({1})" -f $Path, $_.Exception.Message)
    }
}

function New-RdvTestDirectory {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][string]$Name
    )

    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss-fff'
    $path = Join-Path $Root (Join-Path 'build\out\work' ($Name + '-' + $stamp + '-' + $PID))
    New-Item -ItemType Directory -Path $path | Out-Null
    return $path
}
