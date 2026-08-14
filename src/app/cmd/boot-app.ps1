# ---------------------------------------------------------------------------
# boot-app.ps1 -- bootstrap for the practical build's single .cmd.
#
# Same job as src\cmd\boot-common.ps1 / src\v2\cmd\boot-common2.ps1, kept
# separate because those belong to the frozen comparison builds. Differences:
#
#   - defaults are the distribution layout: the CSV subfolder, the ledger xlsx
#     and the execution log all sit next to the .cmd itself
#   - the reference list adds System.IO.Compression and System.Xml (the ledger
#     xlsx is read and written directly, no Excel involved)
#
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

function Rdv-Opt([string[]] $tokens, [string] $name) {
  for ($i = 0; $i -lt $tokens.Count; $i++) {
    if ($tokens[$i] -eq $name -and ($i + 1) -lt $tokens.Count) { return $tokens[$i + 1] }
  }
  return ''
}

function Rdv-DataDir([string[]] $tokens, [string] $here) {
  $cand = New-Object System.Collections.ArrayList
  $skip = $false
  foreach ($t in $tokens) {
    if ($skip) { $skip = $false; continue }
    if ($t -eq '-log' -or $t -eq '-ledger') { $skip = $true; continue }
    if ($t -notlike '-*') { [void]$cand.Add($t) }
  }
  [void]$cand.Add((Join-Path $here 'data'))
  foreach ($c in $cand) {
    if ([string]::IsNullOrEmpty($c)) { continue }
    if (Test-Path -LiteralPath (Join-Path $c 'tableA.csv')) {
      return (Resolve-Path -LiteralPath $c).Path
    }
  }
  return $null
}

function Rdv-Refs() {
  Add-Type -AssemblyName System.Windows.Forms
  Add-Type -AssemblyName System.Drawing
  Add-Type -AssemblyName UIAutomationClient
  Add-Type -AssemblyName UIAutomationTypes
  Add-Type -AssemblyName WindowsBase
  Add-Type -AssemblyName System.IO.Compression
  Add-Type -AssemblyName System.Xml
  return @(
    [System.Diagnostics.Process].Assembly.Location,
    [System.Windows.Forms.Form].Assembly.Location,
    [System.Drawing.Point].Assembly.Location,
    [System.Windows.Automation.AutomationElement].Assembly.Location,
    [System.Windows.Automation.AutomationElementIdentifiers].Assembly.Location,
    [System.Windows.DependencyObject].Assembly.Location,
    [System.IO.Compression.ZipArchive].Assembly.Location,
    [System.Xml.XmlReader].Assembly.Location
  )
}

$RdvHere = Split-Path -Parent $env:RDV_SELF
$RdvTokens = Rdv-Tokens $env:RDV_ARGS
$RdvData = Rdv-DataDir $RdvTokens $RdvHere

$RdvLedger = Rdv-Opt $RdvTokens '-ledger'
if ([string]::IsNullOrEmpty($RdvLedger)) { $RdvLedger = Join-Path $RdvHere 'ReaderDataViewer-Ledger.xlsx' }
$RdvLog = Rdv-Opt $RdvTokens '-log'
if ([string]::IsNullOrEmpty($RdvLog)) { $RdvLog = Join-Path $RdvHere 'ReaderDataViewer.log' }

if ([IntPtr]::Size -ne 8) {
  Rdv-Fail "This build needs 64-bit Windows PowerShell."
}
if (-not $RdvData) {
  Rdv-Fail "tableA.csv / tableB.csv / tableC.csv were not found.`r`nExpected them in the 'data' folder next to the .cmd, or pass the folder:`r`n  $($env:RDV_SELF) <dataDir>"
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

# see boot-csharp.ps1: exit does not evaluate a method call in its argument
$RdvRc = [Rdv3Program]::Run($RdvData, $RdvLedger, $RdvLog, $RdvCompileMs)
exit $RdvRc
