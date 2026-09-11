param([string]$Evidence='', [string]$BaselineRoot='', [switch]$Excel)
$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
if(-not $Evidence){$Evidence=Join-Path $root ('work/review-tests-'+[Guid]::NewGuid().ToString('N'))}
$Evidence=[IO.Path]::GetFullPath($Evidence)
if(Test-Path -LiteralPath $Evidence){throw 'Use a new evidence directory'}
[IO.Directory]::CreateDirectory($Evidence)|Out-Null
$hostExe=Join-Path $PSHOME 'powershell.exe'
function Run([string]$Name,[string]$Program,[string[]]$Arguments){
    $previous=$ErrorActionPreference
    try{$ErrorActionPreference='Continue'; & $Program @Arguments 2>&1 | ForEach-Object {$_.ToString()} | Tee-Object -FilePath (Join-Path $Evidence ($Name+'.txt'));$code=$LASTEXITCODE}
    finally{$ErrorActionPreference=$previous}
    if($code -ne 0){throw ('Failed: '+$Name)}
}
Push-Location -LiteralPath $root
try{
    Run 'core' $hostExe @('-NoProfile','-STA','-ExecutionPolicy','Bypass','-File','tools/Build.ps1','test')
    Run 'archive' $hostExe @('-NoProfile','-STA','-ExecutionPolicy','Bypass','-File','tests/Test-ArchiveSeparation.ps1','-Evidence',(Join-Path $Evidence 'archive-files'))
    Run 'relocation' $hostExe @('-NoProfile','-STA','-ExecutionPolicy','Bypass','-File','tests/Test-Relocation.ps1','-Evidence',(Join-Path $Evidence 'relocation-files'))
    Run 'build' $hostExe @('-NoProfile','-STA','-ExecutionPolicy','Bypass','-File','tests/Test-Build.ps1')
    Run 'boundary' $hostExe @('-NoProfile','-STA','-ExecutionPolicy','Bypass','-File','tests/Test-ReleaseBoundary.ps1','-Evidence',(Join-Path $Evidence 'boundary.json'))
    Run 'samples' 'python' @('tests/test_monthly_samples.py')
    Run 'web' 'python' @('tests/test_web.py')
    Run 'public-records' 'python' @('tests/test_public_results.py')
    if($BaselineRoot){Run 'spec-parity' $hostExe @('-NoProfile','-STA','-ExecutionPolicy','Bypass','-File','tests/Test-SpecParity.ps1','-BaselineRoot',$BaselineRoot,'-Evidence',(Join-Path $Evidence 'spec-parity'))}
    if($Excel){Run 'excel' $hostExe @('-NoProfile','-STA','-ExecutionPolicy','Bypass','-File','tests/Test-ExcelRoundTrip.ps1','-Evidence',(Join-Path $Evidence 'excel-files'))}
}finally{
    try{
        & python (Join-Path $PSScriptRoot 'public_results.py') --path (Join-Path $PSScriptRoot 'results')
        if($LASTEXITCODE -ne 0){throw 'Public result sanitization failed'}
    }finally{Pop-Location}
}
'PASS review regression entry. Native window operations are a separate check; see RELOCATION-20260911.md and REVIEW-20260911.md.'
