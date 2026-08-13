# ============================================================================
# Excel-DNA を nuget.org から取ってきて xll\nuget\ へ展開する。
#
# 必要なのは 2 つだけ。
#   ExcelDna.Integration  … コンパイル時の参照 DLL
#   ExcelDna.AddIn        … 素の loader (ExcelDna64.xll) が入っている
# ExcelDnaPack は使わない。build_xll.ps1 が loader をそのまま置いて、
# 隣に .dna と自前 DLL を並べる形にしている。
#
# nuget.exe も dotnet SDK も要らない。flatcontainer の URL から .nupkg を
# そのまま落とす (.nupkg は zip)。
# ============================================================================
param(
    [string] $Root = "",
    [string] $Version = "1.9.0"
)
$ErrorActionPreference = 'Stop'
if ([string]::IsNullOrEmpty($Root)) { $Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path) }

$nu = Join-Path $Root 'xll\nuget'
New-Item -ItemType Directory -Path $nu -Force | Out-Null

foreach ($id in @('exceldna.integration', 'exceldna.addin')) {
    $url = "https://api.nuget.org/v3-flatcontainer/$id/$Version/$id.$Version.nupkg"
    $zip = Join-Path $nu "$id.zip"
    try {
        Invoke-WebRequest $url -OutFile $zip -TimeoutSec 120 -UseBasicParsing
        Write-Output ("OK  {0} {1}  {2:N0} bytes" -f $id, $Version, (Get-Item $zip).Length)
    }
    catch {
        Write-Output ("NG  {0}: {1}" -f $id, $_.Exception.Message)
        exit 1
    }
    Expand-Archive -LiteralPath $zip -DestinationPath (Join-Path $nu $id) -Force
}

# build_xll.ps1 が探す 2 つが出てきたか確かめる
foreach ($f in @('exceldna.integration\lib\net462\ExcelDna.Integration.dll',
                 'exceldna.addin\tools\net462\ExcelDna64.xll')) {
    $p = Join-Path $nu $f
    if (Test-Path -LiteralPath $p) { Write-Output ("FOUND  " + $f) }
    else { Write-Output ("MISSING  " + $f); exit 1 }
}
