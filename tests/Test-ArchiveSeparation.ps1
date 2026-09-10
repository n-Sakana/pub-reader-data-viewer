param([string]$Root='', [string]$Evidence='')
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.Encoding]::UTF8
if(-not $Root){$Root=Split-Path -Parent $PSScriptRoot}
if(-not $Evidence){$Evidence=Join-Path $Root ('work/archive-separation-'+[Guid]::NewGuid().ToString('N'))}
if(Test-Path -LiteralPath $Evidence){throw 'Use a new evidence directory'}
. (Join-Path $Root 'build/test_support.ps1')
$assemblies=Import-RdvWebView2 -Root $Root
$files=@(Get-RdvSourceFiles -Root $Root)+@(Get-Item -LiteralPath (Join-Path $PSScriptRoot 'ArchiveSeparation.cs'))
$combined=($files|ForEach-Object{[IO.File]::ReadAllText($_.FullName)}) -join [Environment]::NewLine
$pattern='(?m)^\s*using\s+[A-Za-z_][A-Za-z0-9_.]*\s*;\s*$'
$usings=[regex]::Matches($combined,$pattern)|ForEach-Object{$_.Value.Trim()}|Sort-Object -Unique
$source=($usings -join [Environment]::NewLine)+[Environment]::NewLine+[regex]::Replace($combined,$pattern,'')
Add-Type -TypeDefinition $source -ReferencedAssemblies (Get-RdvReferences -WebViewAssemblies $assemblies) -Language CSharp
[ArchiveSeparation]::Run($Root,[IO.Path]::GetFullPath($Evidence))
