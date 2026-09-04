# ============================================================================
# compile_check.ps1 -- compile the product sources exactly as the packer does
# (usings hoisted, one Add-Type under the in-box csc, C# 5), without building
# the distribution. Prints the compiler errors with file-relative line numbers.
#
#   powershell -File build\compile_check.ps1
# ============================================================================
[CmdletBinding()]
param([string] $Root = "")
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
if ([string]::IsNullOrEmpty($Root)) { $Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path) }

. (Join-Path $Root 'build\sources.ps1')

$usings = New-Object System.Collections.Specialized.OrderedDictionary
$bodies = New-Object System.Text.StringBuilder
$map = New-Object System.Collections.ArrayList     # packed line -> (file, line)
foreach ($f in $RdvSources) {
  $p = Join-Path $Root "src\csharp\$f"
  if (-not (Test-Path -LiteralPath $p)) { throw "missing source: $p" }
  $t = [IO.File]::ReadAllText($p, [Text.Encoding]::UTF8)
  if ($t.Contains('@"')) { throw "$f has a verbatim string" }
  $n = 0
  foreach ($line in ($t -split "`r?`n")) {
    $n++
    if ($line -match '^\s*using\s+[A-Za-z_][A-Za-z0-9_.]*\s*;\s*$') {
      $k = $line.Trim(); if (-not $usings.Contains($k)) { $usings.Add($k, $true) }
    } else {
      [void]$bodies.AppendLine($line)
      [void]$map.Add(@($f, $n))
    }
  }
  [void]$bodies.AppendLine()
  [void]$map.Add(@($f, $n + 1))
}
$head = (($usings.Keys | ForEach-Object { $_ }) -join "`r`n") + "`r`n`r`n"
$headLines = ($head -split "`r?`n").Count - 1
$cs = $head + $bodies.ToString()

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.Xml
$refs = @(
  [System.Diagnostics.Process].Assembly.Location,
  [System.Windows.Forms.Form].Assembly.Location,
  [System.Drawing.Point].Assembly.Location,
  [System.Windows.Automation.AutomationElement].Assembly.Location,
  [System.Windows.Automation.AutomationElementIdentifiers].Assembly.Location,
  [System.Windows.DependencyObject].Assembly.Location,
  [System.IO.Compression.ZipArchive].Assembly.Location,
  [System.Xml.XmlReader].Assembly.Location)

$sw = [Diagnostics.Stopwatch]::StartNew()
try {
  Add-Type -TypeDefinition $cs -ReferencedAssemblies $refs -Language CSharp -ErrorAction Stop
} catch {
  $errs = $_.Exception.Message
  Write-Output 'COMPILE FAILED'
  foreach ($m in [regex]::Matches($errs, '\((\d+),(\d+)\)\s*:\s*error\s+(CS\d+):\s*([^\r\n]*)')) {
    $ln = [int]$m.Groups[1].Value - $headLines
    if ($ln -ge 1 -and $ln -le $map.Count) {
      $src = $map[$ln - 1]
      Write-Output ('  {0}({1}): {2}: {3}' -f $src[0], $src[1], $m.Groups[3].Value, $m.Groups[4].Value)
    } else {
      Write-Output ('  (packed line {0}): {1}: {2}' -f $m.Groups[1].Value, $m.Groups[3].Value, $m.Groups[4].Value)
    }
  }
  if ($errs -notmatch 'error CS') { Write-Output $errs }
  exit 1
}
Write-Output ('compile ok ({0} sources, {1:F1} s)' -f $RdvSources.Count, $sw.Elapsed.TotalSeconds)
