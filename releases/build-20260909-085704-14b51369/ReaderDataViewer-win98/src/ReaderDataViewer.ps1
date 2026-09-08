[CmdletBinding(PositionalBinding = $false)]
param(
    [switch]$CompileOnly,
    [switch]$TestCore,
    [switch]$ValidateOnly,
    [switch]$RunUpdate,
    [string]$Config,
    [string]$DataDir,
    [string]$Output,
    [string]$BaselineLedger,
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

    if ('Rdv3Log' -as [type]) { [Rdv3Log]::Feedback($Level,$Message); return }
    # Compilation can fail before Rdv3Log exists. Use the same path, limit,
    # mutex name and rotations for this small bootstrap fallback.
    $roots=@([Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData),[IO.Path]::GetTempPath())
    foreach($root in $roots) {
        $mutex=$null; $owned=$false
        try {
            $logPath=[IO.Path]::GetFullPath((Join-Path $root 'ReaderDataViewer/logs/feedback.log'))
            $line=(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff')+"`tpid=$PID`t$Level`t$Message`r`n"
            $hash=[Security.Cryptography.SHA256]::Create()
            try { $name='Rdv3Log-'+[BitConverter]::ToString($hash.ComputeHash([Text.Encoding]::UTF8.GetBytes($logPath.ToUpperInvariant()))).Replace('-','') }
            finally { $hash.Dispose() }
            $mutex=New-Object Threading.Mutex($false,$name)
            try { $owned=$mutex.WaitOne(1000) } catch [Threading.AbandonedMutexException] { $owned=$true }
            if(-not $owned) { throw 'Log is busy' }
            [IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($logPath)) | Out-Null
            if([Text.Encoding]::UTF8.GetByteCount($line) -gt 4194304) { $line=$line.Substring(0,[Math]::Min($line.Length,1048476))+"`r`n[LOG ENTRY TRUNCATED: exceeded 4 MiB]`r`n" }
            $size=if([IO.File]::Exists($logPath)) {(New-Object IO.FileInfo($logPath)).Length} else {0}
            if($size+[Text.Encoding]::UTF8.GetByteCount($line) -gt 4194304) {
                for($i=3;$i -ge 1;$i--) {
                    $older=$logPath+'.'+$i
                    $newer=if($i -eq 1) {$logPath} else {$logPath+'.'+($i-1)}
                    if([IO.File]::Exists($older)) { [IO.File]::Delete($older) }
                    if([IO.File]::Exists($newer)) { [IO.File]::Move($newer,$older) }
                }
            }
            $file=New-Object IO.FileStream($logPath,[IO.FileMode]::Append,[IO.FileAccess]::Write,[IO.FileShare]::Read)
            try { $bytes=[Text.Encoding]::UTF8.GetBytes($line); $file.Write($bytes,0,$bytes.Length); $file.Flush($true) }
            finally { $file.Dispose() }
            return
        } catch { [Console]::Error.WriteLine('LOG WRITE FAILED '+$root+': '+$_.Exception.Message) }
        finally { if($owned) { $mutex.ReleaseMutex() }; if($null -ne $mutex) { $mutex.Dispose() } }
    }
}

$exitCode=$startupFailureExitCode
$context='app='+(Split-Path -Parent $PSScriptRoot)+' config='+$Config+' data='+$DataDir+' output='+$Output+' validate='+$ValidateOnly+' update='+$RunUpdate
Write-ReaderLauncherLog 'BEGIN BOOTSTRAP' $context
[Console]::Error.WriteLine('LOG '+(Join-Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)) 'ReaderDataViewer/logs/feedback.log'))
try {
    if ($ReaderArguments -and $ReaderArguments.Count -gt 0) {
        $detail='Unknown arguments: ' + ($ReaderArguments -join ' ')
        Write-ReaderLauncherLog 'ERROR' $detail
        [Console]::Error.WriteLine($detail)
        $exitCode=2
        exit $exitCode
    }
    # PowerShell 7 (pwsh) runs on .NET Core, where Add-Type cannot compile the
    # WPF/WebView2 sources (System.Drawing.Color, PresentationFramework). Say so
    # before compiling, instead of leaving a compiler error to explain it.
    if ($PSVersionTable.PSEdition -eq 'Core') {
        $detail='Windows PowerShell 5.1 (powershell.exe) is required. PowerShell ' + $PSVersionTable.PSVersion +
            ' (pwsh) cannot compile the WPF/WebView2 sources. Run: powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File "' +
            $PSCommandPath + '" ... / Windows PowerShell 5.1 (powershell.exe) で実行してください。pwsh では起動できません。'
        Write-ReaderLauncherLog 'ERROR' $detail
        [Console]::Error.WriteLine('Reader Data Viewer: ' + $detail)
        $exitCode=3
        exit $exitCode
    }
    # This script lives in src\; the application root -- the folder holding
    # settings.json, src\, web\, lib\ and data\ -- is its parent.
    $baseDirectory = Split-Path -Parent $PSScriptRoot
    $headless = $ValidateOnly -or $RunUpdate
    if (($ValidateOnly -and $RunUpdate) -or ($headless -and ($TestCore -or $CompileOnly))) {
        throw 'Choose one mode: -ValidateOnly, -RunUpdate, -CompileOnly or -TestCore.'
    }
    if (-not $headless -and ($Config -or $DataDir -or $Output -or $BaselineLedger)) {
        throw '-Config, -DataDir, -Output and -BaselineLedger require -ValidateOnly or -RunUpdate.'
    }
    if ($RunUpdate -and -not $Output) { throw '-RunUpdate requires -Output <new.json>.' }
    if ($ValidateOnly -and ($Output -or $BaselineLedger)) { throw '-Output and -BaselineLedger require -RunUpdate.' }
    $sourceDirectory = Join-Path $baseDirectory 'src'
    $libraryDirectory = Join-Path $baseDirectory 'lib'
    Write-ReaderLauncherLog 'PHASE' 'loading assemblies and compiling application sources'

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

    if ($TestCore) {
        $sourceFiles += Get-Item -LiteralPath (Join-Path $baseDirectory 'tests\RegressionTests.cs')
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

    [Rdv3Log]::Begin('launcher',$context)

    if ($TestCore) {
        $exitCode = [Rdv3RegressionTests]::Run($baseDirectory)
        exit $exitCode
    }
    if ($CompileOnly) {
        [Console]::WriteLine('PASS: all application C# sources compiled. No window or ledger was opened.')
        $exitCode=0
        exit $exitCode
    }
    if ($headless) {
        if (-not $Config) { $Config = Join-Path $baseDirectory 'settings.json' }
        $Config = [IO.Path]::GetFullPath($Config)
        if ($DataDir) { $DataDir = [IO.Path]::GetFullPath($DataDir) }
        if ($Output) { $Output = [IO.Path]::GetFullPath($Output) }
        if ($BaselineLedger) { $BaselineLedger = [IO.Path]::GetFullPath($BaselineLedger) }
        $exitCode = [Rdv3Headless]::Run($baseDirectory, $Config, $DataDir, $RunUpdate.IsPresent, $Output, $BaselineLedger)
        exit $exitCode
    }
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
    $exitCode=$startupFailureExitCode
    exit $exitCode
}
finally {
    if ('Rdv3Log' -as [type]) { [Rdv3Log]::End($exitCode) }
    else { Write-ReaderLauncherLog 'END' ('exit='+$exitCode) }
}
