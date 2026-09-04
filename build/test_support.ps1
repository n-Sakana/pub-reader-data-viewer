[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

function Import-RdvProduct {
    param([Parameter(Mandatory = $true)][string]$Root)

    $sourceDirectory = Join-Path $Root 'src'
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
        [Reflection.Assembly]::Load([IO.File]::ReadAllBytes($assemblyPath)) | Out-Null
    }

    $sourceFiles = @(Get-ChildItem -LiteralPath $sourceDirectory -Filter '*.cs' -File |
        Sort-Object -Property Name)
    if ($sourceFiles.Count -eq 0) {
        throw "No C# source files were found in: $sourceDirectory"
    }

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

    $references = @(
        [System.Windows.Window].Assembly.Location
        [System.Windows.UIElement].Assembly.Location
        [System.Windows.DependencyObject].Assembly.Location
        [System.Xaml.XamlReader].Assembly.Location
        [System.Windows.Automation.AutomationElement].Assembly.Location
        [System.Windows.Automation.ControlType].Assembly.Location
        [System.IO.Compression.ZipArchive].Assembly.Location
        [System.Xml.XmlDocument].Assembly.Location
        'System.Drawing'
        $webViewAssemblies[0]
        $webViewAssemblies[1]
    )

    Add-Type -TypeDefinition $source `
        -ReferencedAssemblies $references `
        -Language CSharp

    Write-Output ("compile ok ({0} sources)" -f $sourceFiles.Count)
}

function New-RdvTestDirectory {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][string]$Name
    )

    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss-fff'
    $path = Join-Path $Root (Join-Path 'work' ($Name + '-' + $stamp + '-' + $PID))
    New-Item -ItemType Directory -Path $path | Out-Null
    return $path
}
