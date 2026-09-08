param([Parameter(Mandatory=$true)][string]$OutputDirectory)
$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
[IO.Directory]::CreateDirectory([IO.Path]::GetFullPath($OutputDirectory)) | Out-Null
$combined=[IO.File]::ReadAllText((Join-Path $root 'src/Rdv3Log.cs'))+"`n"+[IO.File]::ReadAllText((Join-Path $PSScriptRoot 'LoggingFaults.cs'))
$pattern='(?m)^\s*using\s+[A-Za-z_][A-Za-z0-9_.]*\s*;\s*$'
$usings=[regex]::Matches($combined,$pattern) | ForEach-Object {$_.Value.Trim()} | Sort-Object -Unique
$source=($usings -join "`n")+"`n"+[regex]::Replace($combined,$pattern,'')
Add-Type -TypeDefinition $source -OutputType ConsoleApplication -OutputAssembly (Join-Path $OutputDirectory 'LoggingFaults.exe')
Write-Output (Join-Path $OutputDirectory 'LoggingFaults.exe')
