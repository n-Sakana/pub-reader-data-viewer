# ============================================================================
# Excel-DNA の XLL を組み立てる。
#
# ExcelDnaPack は使わない。Excel-DNA が配っている素の loader (ExcelDna64.xll) を
# 名前だけ変えて置き、隣に .dna と自前の DLL を並べる形にする。
# 中身が見える分、何が読み込まれているか追いやすい。
#
# Excel が x64 なので loader も 64bit 版を使う。取り違えると Excel は
# 「有効な アドイン ではありません」とだけ言って理由を教えてくれない。
# ============================================================================
param([string] $Root = "")
$ErrorActionPreference = 'Stop'
if ([string]::IsNullOrEmpty($Root)) { $Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path) }

$xll = Join-Path $Root 'xll'
$nu  = Join-Path $xll 'nuget'
$out = Join-Path $Root 'prebuilt'
New-Item -ItemType Directory -Path $out -Force | Out-Null

$integration = Join-Path $nu 'exceldna.integration\lib\net462\ExcelDna.Integration.dll'
$loader      = Join-Path $nu 'exceldna.addin\tools\net462\ExcelDna64.xll'
foreach ($f in @($integration, $loader)) {
    if (-not (Test-Path -LiteralPath $f)) { Write-Output "MISSING  $f"; exit 1 }
}

$csc = Join-Path $env:SystemRoot 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (-not (Test-Path -LiteralPath $csc)) {
    $csc = Join-Path $env:SystemRoot 'Microsoft.NET\Framework\v4.0.30319\csc.exe'
}
if (-not (Test-Path -LiteralPath $csc)) { Write-Output "MISSING  csc.exe"; exit 1 }

$dll = Join-Path $out 'ZbXllLib.dll'
$log = Join-Path $out 'xll_build.log'

$sw = [Diagnostics.Stopwatch]::StartNew()
$args = @(
    '/nologo', '/target:library', '/platform:x64', '/optimize+',
    "/reference:$integration",
    "/out:$dll",
    (Join-Path $xll 'ZbXll.cs')
)
$output = & $csc @args 2>&1
$output | Out-File -LiteralPath $log -Encoding utf8

if (-not (Test-Path -LiteralPath $dll)) {
    $output | ForEach-Object { Write-Output $_ }
    Write-Output "BUILD_FAILED  see $log"
    exit 1
}

# loader と参照 DLL を並べる
Copy-Item $loader (Join-Path $out 'ZbXll64.xll') -Force
Copy-Item $integration (Join-Path $out 'ExcelDna.Integration.dll') -Force

# .dna は loader と同じ base name でなければ読まれない
$dna = @'
<DnaLibrary Name="ZbXll" RuntimeVersion="v4.0">
  <ExternalLibrary Path="ZbXllLib.dll" ExplicitExports="false" LoadFromBytes="false" />
</DnaLibrary>
'@
Set-Content -LiteralPath (Join-Path $out 'ZbXll64.dna') -Value $dna -Encoding UTF8

Write-Output ("BUILD_OK  {0}  {1} bytes  {2} ms" -f $dll, (Get-Item -LiteralPath $dll).Length, $sw.ElapsedMilliseconds)
Write-Output ("XLL       " + (Join-Path $out 'ZbXll64.xll'))
