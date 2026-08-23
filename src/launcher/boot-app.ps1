# ---------------------------------------------------------------------------
# boot-app.ps1 -- bootstrap for the practical build's single .cmd.
#
# Product bootstrap. Earlier variants are frozen under archive\comparisons.
# It compiles the embedded C# with the in-box csc and hands over to
# Rdv3Program.Run. Everything about the configuration -- reading settings.json,
# checking it, resolving the paths in it -- happens in the C#; this script only
# forwards what the command line said:
#
#   ReaderDataViewer.cmd [dataDir] [-config <settings.json>] [-ledger <file.xlsx>] [-log <file>]
#
# Relative paths are relative to the folder holding the .cmd. The settings file
# defaults to settings.json in that folder.
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

# the first token that is not an option or an option's value: the data folder
function Rdv-Positional([string[]] $tokens) {
  $skip = $false
  foreach ($t in $tokens) {
    if ($skip) { $skip = $false; continue }
    if ($t -eq '-log' -or $t -eq '-ledger' -or $t -eq '-config') { $skip = $true; continue }
    if ($t -notlike '-*') { return $t }
  }
  return ''
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

$RdvConfig = Rdv-Opt $RdvTokens '-config'
if ([string]::IsNullOrEmpty($RdvConfig)) { $RdvConfig = Join-Path $RdvHere 'settings.json' }
elseif (-not [IO.Path]::IsPathRooted($RdvConfig)) { $RdvConfig = Join-Path $RdvHere $RdvConfig }
$RdvDataArg = Rdv-Positional $RdvTokens
$RdvLedgerArg = Rdv-Opt $RdvTokens '-ledger'
$RdvLogArg = Rdv-Opt $RdvTokens '-log'

if ([IntPtr]::Size -ne 8) {
  Rdv-Fail "This build needs 64-bit Windows PowerShell."
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
$RdvRc = [Rdv3Program]::Run($RdvConfig, $RdvHere, $RdvDataArg, $RdvLedgerArg, $RdvLogArg, $RdvCompileMs)
exit $RdvRc
