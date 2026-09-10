# Exercise package boundaries in an isolated source tree containing deliberate residue.
[CmdletBinding()]
param([string]$Evidence = '')
Set-StrictMode -Version 2
$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
$hostExe=(Get-Process -Id $PID).Path
$temp=Join-Path ([IO.Path]::GetTempPath()) ('rdv-release-boundary-'+[Guid]::NewGuid().ToString('N'))
$copy=Join-Path $temp 'source'
$utf8=New-Object Text.UTF8Encoding($false)
$results=New-Object Collections.Generic.List[object]
function Check([bool]$Condition,[string]$Message){if(-not $Condition){throw $Message}}
function Test([string]$Name,[scriptblock]$Body){
    try{& $Body;$results.Add(@{name=$Name;passed=$true});Write-Output ('PASS '+$Name)}
    catch{$results.Add(@{name=$Name;passed=$false;error=$_.Exception.Message});Write-Output ('FAIL '+$Name+': '+$_.Exception.Message)}
}
function Put([string]$Name,[string]$Text){
    $path=Join-Path $copy $Name
    [IO.Directory]::CreateDirectory((Split-Path -Parent $path))|Out-Null
    [IO.File]::WriteAllText($path,$Text,$utf8)
}
function Copy-Source([string]$Name){
    $target=Join-Path $copy $Name
    [IO.Directory]::CreateDirectory((Split-Path -Parent $target))|Out-Null
    Copy-Item -LiteralPath (Join-Path $root $Name) -Destination $target
}
function Build([string]$Kind,[string]$Data,[int]$Expected=0,[bool]$Legacy=$false){
    $out=Join-Path $temp $Kind
    $script=Join-Path $copy 'tools/Build.ps1'
    $args=@('package')
    if($Legacy){$script=Join-Path $copy 'build/build_dist.ps1';$args=@('-Root',$copy)}
    & $hostExe -NoProfile -ExecutionPolicy Bypass -File $script @args -Data $Data -Format both -OutputRoot $out -SkipValidation
    Check ($LASTEXITCODE -eq $Expected) ('Unexpected exit: '+$LASTEXITCODE)
    $built=@(Get-ChildItem -LiteralPath $out -Directory -Filter 'build-*' -ErrorAction SilentlyContinue)
    if($Expected -ne 0){Check ($built.Count -eq 0) 'Failed build left a published output';return}
    Check ($built.Count -eq 1) 'Expected one product'
    return Join-Path $built[0].FullName $spec.packageName
}
try{
    [IO.Directory]::CreateDirectory($copy)|Out-Null
    $spec=Get-Content -LiteralPath (Join-Path $root 'tools/package-files.json') -Raw -Encoding UTF8|ConvertFrom-Json
    foreach($file in $spec.files){Copy-Source $file}
    foreach($file in @('tools/Build.ps1','tools/package-files.json','build/build_dist.ps1','configs/sample/settings.json')){Copy-Source $file}
    foreach($dir in @('tests/fixtures/sample-v4','samples')){
        foreach($file in (Get-ChildItem -LiteralPath (Join-Path $root $dir) -Recurse -File)){
            Copy-Source $file.FullName.Substring($root.Length+1)
        }
    }
    Put 'settings.json' '{"private-setting-marker":"NOT FOR DISTRIBUTION"}'
    Put 'data/tableA.csv' 'PRIVATE DATA MARKER'
    Put 'src/leftover-win32.cs' 'DEVELOPMENT RESIDUE'
    Put 'web/review-note.md' 'DEVELOPMENT RESIDUE'
    Put 'lib/Win32/old.dll' 'DEVELOPMENT RESIDUE'
    Put 'samples/next-period/workbench.txt' 'DEVELOPMENT RESIDUE'
    Put 'configs/production/settings.json' '{"private-setting-marker":"NOT FOR DISTRIBUTION"}'
    $expected=(Get-FileHash -LiteralPath (Join-Path $root 'configs/sample/settings.json')).Hash
    Test 'sample-excludes-residue-and-local-settings' {
        $package=@(Build 'sample' 'sample')[-1]
        Check ((Get-FileHash -LiteralPath (Join-Path $package 'settings.json')).Hash -eq $expected) 'Local settings leaked'
        $files=@(Get-ChildItem -LiteralPath $package -Recurse -File)
        Check (@($files|Where-Object {$_.Name -in @('leftover-win32.cs','review-note.md','old.dll','workbench.txt','tableA.csv')}).Count -eq 0) 'Unlisted files leaked'
        Check (-not(Test-Path -LiteralPath (Join-Path $package 'configs'))) 'Private configuration directory leaked'
        $release=Get-Content -LiteralPath (Join-Path $package 'manual/release.json') -Raw -Encoding UTF8|ConvertFrom-Json
        Check ($null -eq $release.sourceCommit) 'Exported source tree invented a commit'
    }
    Test 'no-data-also-uses-public-template' {
        $package=@(Build 'none' 'none')[-1]
        Check ((Get-FileHash -LiteralPath (Join-Path $package 'settings.json')).Hash -eq $expected) 'Data none leaked local config'
        Check (@(Get-ChildItem -LiteralPath (Join-Path $package 'data') -File).Count -eq 0) 'Private input was included'
        Check (-not(Test-Path -LiteralPath (Join-Path $package 'samples'))) 'Data none included examples'
    }
    Test 'legacy-builder-uses-identical-boundaries' {
        $package=@(Build 'legacy' 'sample' 0 $true)[-1]
        Check ((Get-FileHash -LiteralPath (Join-Path $package 'settings.json')).Hash -eq $expected) 'Legacy path copied local config'
        Check (-not(Test-Path -LiteralPath (Join-Path $package 'src/leftover-win32.cs'))) 'Legacy path copied residue'
        Check (-not(Test-Path -LiteralPath (Join-Path $copy 'build/out'))) 'Legacy output route still ran'
    }
    Test 'missing-required-runtime-file-refuses-publication' {
        $path=Join-Path $copy 'web/app.js'
        Remove-Item -LiteralPath $path
        Build 'missing' 'sample' 1|Out-Null
        Copy-Source 'web/app.js'
    }
    Test 'manifest-cannot-escape-package' {
        $spec.files += '../outside.txt'
        Put 'tools/package-files.json' ($spec|ConvertTo-Json -Depth 5)
        Build 'escape' 'none' 1|Out-Null
        Check (-not(Test-Path -LiteralPath (Join-Path $temp 'outside.txt'))) 'Manifest escaped package'
    }
}finally{
    $resolved=[IO.Path]::GetFullPath($temp)
    $owner=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd([IO.Path]::DirectorySeparatorChar)
    if(-not $resolved.StartsWith($owner+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase) -or [IO.Path]::GetFileName($resolved) -notlike 'rdv-release-boundary-*'){throw 'Unsafe cleanup'}
    if(Test-Path -LiteralPath $resolved){Remove-Item -LiteralPath $resolved -Recurse -Force}
}
$failed=@($results|Where-Object {-not $_.passed}).Count
if($Evidence){
    [IO.Directory]::CreateDirectory((Split-Path -Parent ([IO.Path]::GetFullPath($Evidence))))|Out-Null
    [IO.File]::WriteAllText($Evidence,(@{tests=$results.ToArray();failed=$failed}|ConvertTo-Json -Depth 6),$utf8)
}
if($failed){exit 1}
