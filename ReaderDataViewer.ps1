[CmdletBinding(PositionalBinding = $false)]
param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$ReaderArguments
)

Set-StrictMode -Version 2
$ErrorActionPreference = 'Stop'
$startupFailureExitCode = 3

function Write-ReaderLauncherLog {
    param(
        [string]$Level,
        [string]$Message
    )

    try {
        $localData = [Environment]::GetFolderPath(
            [Environment+SpecialFolder]::LocalApplicationData)
        if ([string]::IsNullOrWhiteSpace($localData)) {
            return
        }

        $logDirectory = Join-Path $localData 'ReaderDataViewer\logs'
        if (-not (Test-Path -LiteralPath $logDirectory -PathType Container)) {
            New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
        }

        $logPath = Join-Path $logDirectory (
            'reader-data-viewer_' + (Get-Date -Format 'yyyyMMdd') + '.log')
        $line = '[' + (Get-Date -Format 'HH:mm:ss') + '] ' +
            '[' + $Level + '] ' + $Message
        [IO.File]::AppendAllText(
            $logPath,
            $line + [Environment]::NewLine,
            (New-Object Text.UTF8Encoding($false)))
    }
    catch {
    }
}

try {
    $baseDirectory = $PSScriptRoot
    $sourceDirectory = Join-Path $baseDirectory 'src'
    $libraryDirectory = Join-Path $baseDirectory 'lib'

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

        [Reflection.Assembly]::Load(
            [IO.File]::ReadAllBytes($assemblyPath)) | Out-Null
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

    $exitCode = [ReaderDataViewer.App]::Run($baseDirectory)
    exit $exitCode
}
catch {
    $location = ''
    if ($null -ne $_.InvocationInfo) {
        $location = ' (' + $_.InvocationInfo.ScriptName + ':' +
            $_.InvocationInfo.ScriptLineNumber + ')'
    }
    $detail = 'launcher failed' + $location + ' ' + $_.Exception.ToString()
    Write-ReaderLauncherLog 'ERROR' $detail
    [Console]::Error.WriteLine('Reader Data Viewer: ' + $detail)
    exit $startupFailureExitCode
}
