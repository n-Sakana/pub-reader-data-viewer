param([Parameter(Mandatory=$true)][string]$Build,[string]$OutputRoot='')
$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
$Build=(Resolve-Path -LiteralPath $Build).Path
if(-not $OutputRoot){$OutputRoot=Join-Path $root 'build/packages'}
$OutputRoot=[IO.Path]::GetFullPath($OutputRoot)
$name='review-'+(Get-Date -Format yyyyMMdd-HHmmss)+'-'+[Guid]::NewGuid().ToString('N').Substring(0,8)
$destination=Join-Path $OutputRoot $name
$source=Join-Path $destination 'source'
[IO.Directory]::CreateDirectory($source)|Out-Null
$spec=Get-Content -LiteralPath (Join-Path $root 'tools/package-files.json') -Raw -Encoding UTF8|ConvertFrom-Json
$files=@($spec.files)+@('build.bat','tools/Build.ps1','tools/package-files.json','tools/Build-Review.ps1','configs/sample/settings.json')
$tracked=@(& git -C $root -c core.quotepath=false ls-files -- tests build docs)
if($LASTEXITCODE -ne 0){throw 'Git is required to select reviewed source files'}
$files+=$tracked
$files+=@('build/gen_monthly_samples.py','tests/ReviewDesktop.cs','tests/ReviewDesktop.ps1','tests/Test-ArchiveSeparation.ps1','tests/Test-ExcelRoundTrip.ps1','tests/Test-SpecParity.ps1','tests/Test-Review.ps1','tests/test_monthly_samples.py','tests/REVIEW-20260911.md')
$files+=@('tests/Relocation.cs','tests/Test-Relocation.ps1')
$files+=@('tests/public_results.py','tests/test_public_results.py')
$files+=@('tests/RELOCATION-20260911.md')
foreach($directory in @('samples/current/data','samples/current/追加CSVデータ/5月分（3件追加）')){
    $csvs=@(Get-ChildItem -LiteralPath (Join-Path $root $directory) -Filter '*.csv' -File)
    if($csvs.Count -ne 4){throw ('Expected four published CSVs: '+$directory)}
    $files+=@($csvs|ForEach-Object {$directory+'/'+$_.Name})
}
foreach($relative in @($files|Sort-Object -Unique)){
    $path=[IO.Path]::GetFullPath((Join-Path $root $relative))
    if(-not $path.StartsWith($root+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){throw 'Source escaped repository'}
    $target=Join-Path $source $relative
    [IO.Directory]::CreateDirectory((Split-Path -Parent $target))|Out-Null
    Copy-Item -LiteralPath $path -Destination $target
}
Copy-Item -LiteralPath (Join-Path $root 'configs/sample/settings.json') -Destination (Join-Path $source 'settings.json')
& python (Join-Path $root 'tests/public_results.py') --path (Join-Path $source 'tests/results')
if($LASTEXITCODE -ne 0){throw 'Could not sanitize public review results'}
& python (Join-Path $root 'tests/public_results.py') --path (Join-Path $source 'tests/results') --check
if($LASTEXITCODE -ne 0){throw 'User profile paths remain in review results'}
foreach($required in @('build-summary.json','package-manifest.json')){
    if(-not(Test-Path -LiteralPath (Join-Path $Build $required))){throw ('Missing build evidence: '+$required)}
}
Copy-Item -LiteralPath $Build -Destination (Join-Path $destination 'distribution') -Recurse
$entries=@(Get-ChildItem -LiteralPath $destination -File -Recurse|Sort-Object FullName|ForEach-Object{
    [ordered]@{path=$_.FullName.Substring($destination.Length+1).Replace('\','/');bytes=$_.Length;sha256=(Get-FileHash -LiteralPath $_.FullName).Hash.ToLowerInvariant()}
})
$manifest=[ordered]@{schema=1;sourceBaseline='3ca9434';priorReviewBaseline='ec6aefc';files=$entries;testEntry='source/tests/Test-Review.ps1';report='source/tests/RELOCATION-20260911.md'}
[IO.File]::WriteAllText((Join-Path $destination 'review-manifest.json'),($manifest|ConvertTo-Json -Depth 6),[Text.UTF8Encoding]::new($false))
Add-Type -AssemblyName System.IO.Compression,System.IO.Compression.FileSystem
$zipPath=$destination+'.zip'
$zip=[IO.Compression.ZipFile]::Open($zipPath,[IO.Compression.ZipArchiveMode]::Create)
try{
    foreach($file in Get-ChildItem -LiteralPath $destination -File -Recurse|Sort-Object FullName){
        $relative=$file.FullName.Substring($destination.Length+1).Replace('\','/')
        [IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip,$file.FullName,$name+'/'+$relative,[IO.Compression.CompressionLevel]::Optimal)|Out-Null
    }
}finally{$zip.Dispose()}
Write-Output ('Review: '+$destination)
Write-Output ('ZIP: '+$zipPath)
