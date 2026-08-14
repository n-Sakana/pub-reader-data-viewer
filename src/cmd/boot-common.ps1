# ---------------------------------------------------------------------------
# boot-common.ps1 -- shared bootstrap, spliced into both .cmd files.
# ASCII only: this text lives in a file cmd.exe also reads, and the console it
# prints to is whatever code page the machine happens to be using.
# ---------------------------------------------------------------------------
$ErrorActionPreference = 'Stop'

function Rdv-Tokens([string] $line) {
  $out = @()
  if ([string]::IsNullOrEmpty($line)) { return $out }
  foreach ($m in [regex]::Matches($line, '"[^"]*"|\S+')) {
    $out += $m.Value.Trim('"')
  }
  return $out
}

function Rdv-Fail([string] $msg) {
  try {
    Add-Type -AssemblyName System.Windows.Forms
    [void][System.Windows.Forms.MessageBox]::Show($msg, 'Reader Data Viewer', 0, 16)
  } catch {
    Write-Host $msg
  }
  exit 1
}

function Rdv-DataDir([string[]] $tokens, [string] $here) {
  $cand = New-Object System.Collections.ArrayList
  foreach ($t in $tokens) { if ($t -notlike '-*') { [void]$cand.Add($t) } }
  if ($env:RDV_DATA) { [void]$cand.Add($env:RDV_DATA) }
  [void]$cand.Add((Join-Path (Split-Path -Parent $here) 'data'))
  [void]$cand.Add((Join-Path $here 'data'))
  foreach ($c in $cand) {
    if ([string]::IsNullOrEmpty($c)) { continue }
    if (Test-Path -LiteralPath (Join-Path $c 'tableA.csv')) {
      return (Resolve-Path -LiteralPath $c).Path
    }
  }
  return $null
}

function Rdv-LogPath([string[]] $tokens) {
  for ($i = 0; $i -lt $tokens.Count; $i++) {
    if ($tokens[$i] -eq '-log' -and ($i + 1) -lt $tokens.Count) { return $tokens[$i + 1] }
  }
  if ($env:RDV_LOG) { return $env:RDV_LOG }
  return ''
}

# every assembly the embedded C# needs, resolved to a real path so csc can find
# it: UIAutomationClient does not sit in the compiler's own directory
function Rdv-Refs() {
  Add-Type -AssemblyName System.Windows.Forms
  Add-Type -AssemblyName System.Drawing
  Add-Type -AssemblyName UIAutomationClient
  Add-Type -AssemblyName UIAutomationTypes
  Add-Type -AssemblyName WindowsBase
  return @(
    [System.Diagnostics.Process].Assembly.Location,
    [System.Windows.Forms.Form].Assembly.Location,
    [System.Drawing.Point].Assembly.Location,
    [System.Windows.Automation.AutomationElement].Assembly.Location,
    [System.Windows.Automation.AutomationElementIdentifiers].Assembly.Location,
    [System.Windows.DependencyObject].Assembly.Location
  )
}

$RdvHere = Split-Path -Parent $env:RDV_SELF
$RdvTokens = Rdv-Tokens $env:RDV_ARGS
$RdvData = Rdv-DataDir $RdvTokens $RdvHere
$RdvLog = Rdv-LogPath $RdvTokens

if ([IntPtr]::Size -ne 8) {
  Rdv-Fail "This build needs 64-bit Windows PowerShell: 1,000,000 rows x 3 tables will not fit in a 32-bit process."
}
if (-not $RdvData) {
  Rdv-Fail "tableA.csv / tableB.csv / tableC.csv were not found. Generate them first:`r`n  powershell -File build\gen_data.ps1`r`nor pass the data directory:`r`n  $($env:RDV_SELF) <dataDir>"
}

$RdvRefs = Rdv-Refs
$RdvSw = [Diagnostics.Stopwatch]::StartNew()
try {
  Add-Type -TypeDefinition $global:RdvCs -ReferencedAssemblies $RdvRefs -Language CSharp
} catch {
  Rdv-Fail ("The embedded C# did not compile:`r`n" + $_.Exception.Message)
}
$RdvSw.Stop()
$RdvCompileMs = $RdvSw.Elapsed.TotalMilliseconds
