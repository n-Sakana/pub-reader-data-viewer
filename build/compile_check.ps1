# ============================================================================
# compile_check.ps1 -- compile the product sources exactly as ReaderDataViewer.ps1
# compiles them (every src\*.cs in name order, usings hoisted, one Add-Type
# under the in-box csc, C# 5), without starting a window or building a
# distribution. Compiler errors are printed with file-relative line numbers.
#
#   powershell -File build\compile_check.ps1
#
# The assembly list and the source list come from build\test_support.ps1, so
# this check and the tests compile the same thing. The concatenation is done
# here rather than through Import-RdvProduct only because this script keeps a
# packed-line -> (file, line) map to translate the errors back.
# ============================================================================
[CmdletBinding()]
param([string] $Root = "")
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
if ([string]::IsNullOrEmpty($Root)) { $Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path) }

. (Join-Path $Root 'build\test_support.ps1')

$webViewAssemblies = Import-RdvWebView2 -Root $Root
$sourceFiles = Get-RdvSourceFiles -Root $Root

$usings = New-Object System.Collections.Specialized.OrderedDictionary
$bodies = New-Object System.Text.StringBuilder
$map = New-Object System.Collections.ArrayList     # packed line -> (file, line)
foreach ($file in $sourceFiles) {
  $text = [IO.File]::ReadAllText($file.FullName, [Text.Encoding]::UTF8)
  $n = 0
  foreach ($line in ($text -split "`r?`n")) {
    $n++
    if ($line -match '^\s*using\s+[A-Za-z_][A-Za-z0-9_.]*\s*;\s*$') {
      $k = $line.Trim(); if (-not $usings.Contains($k)) { $usings.Add($k, $true) }
    } else {
      [void]$bodies.AppendLine($line)
      [void]$map.Add(@($file.Name, $n))
    }
  }
  [void]$bodies.AppendLine()
  [void]$map.Add(@($file.Name, $n + 1))
}
$head = (($usings.Keys | ForEach-Object { $_ }) -join "`r`n") + "`r`n`r`n"
$headLines = ($head -split "`r?`n").Count - 1
$cs = $head + $bodies.ToString()

$sw = [Diagnostics.Stopwatch]::StartNew()
try {
  Add-Type -TypeDefinition $cs `
    -ReferencedAssemblies (Get-RdvReferences -WebViewAssemblies $webViewAssemblies) `
    -Language CSharp -ErrorAction Stop
} catch {
  $errs = $_.Exception.Message
  Write-Output 'COMPILE FAILED'
  foreach ($m in [regex]::Matches($errs, '\((\d+),(\d+)\)\s*:\s*error\s+(CS\d+):\s*([^\r\n]*)')) {
    $ln = [int]$m.Groups[1].Value - $headLines
    if ($ln -ge 1 -and $ln -le $map.Count) {
      $src = $map[$ln - 1]
      Write-Output ('  src\{0}({1}): {2}: {3}' -f $src[0], $src[1], $m.Groups[3].Value, $m.Groups[4].Value)
    } else {
      Write-Output ('  (packed line {0}): {1}: {2}' -f $m.Groups[1].Value, $m.Groups[3].Value, $m.Groups[4].Value)
    }
  }
  if ($errs -notmatch 'error CS') { Write-Output $errs }
  exit 1
}
Write-Output ('compile ok ({0} sources, {1:F1} s)' -f $sourceFiles.Count, $sw.Elapsed.TotalSeconds)
