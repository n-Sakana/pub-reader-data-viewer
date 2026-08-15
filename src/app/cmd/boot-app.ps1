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

# The console window belongs to the launcher, not to the app: hide it as early
# as PowerShell can. Everything the operator needs to see is on the app's own
# window, and a start-up failure still comes up as a message box (Rdv-Fail).
try {
  Add-Type -Name RdvCon -Namespace Rdv -MemberDefinition '
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int n);'
  $RdvConH = [Rdv.RdvCon]::GetConsoleWindow()
  if ($RdvConH -ne [IntPtr]::Zero) { [void][Rdv.RdvCon]::ShowWindow($RdvConH, 0) }
} catch { }

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

function Rdv-DataDir([string[]] $tokens, [string] $here, [string] $fromJson) {
  $cand = New-Object System.Collections.ArrayList
  $skip = $false
  foreach ($t in $tokens) {
    if ($skip) { $skip = $false; continue }
    if ($t -eq '-log' -or $t -eq '-ledger' -or $t -eq '-config') { $skip = $true; continue }
    if ($t -notlike '-*') { [void]$cand.Add($t) }
  }
  if (-not [string]::IsNullOrEmpty($fromJson)) { [void]$cand.Add($fromJson) }
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

# The settings file is the second half of the distribution. It is read here as
# well as in the app, because the preflight below has to know where the CSVs
# are before anything is compiled. Order everywhere: command line, then the
# settings file, then the built-in default.
$RdvConfig = Rdv-Opt $RdvTokens '-config'
if ([string]::IsNullOrEmpty($RdvConfig)) { $RdvConfig = Join-Path $RdvHere 'ReaderDataViewer.json' }
$RdvJson = $null
if (Test-Path -LiteralPath $RdvConfig) {
  try {
    $RdvRaw = [IO.File]::ReadAllText($RdvConfig, [Text.Encoding]::UTF8)
    $RdvRaw = [regex]::Replace($RdvRaw, '(?m)^\s*//.*$', '')
    $RdvJson = $RdvRaw | ConvertFrom-Json
  } catch { $RdvJson = $null }
}
function Rdv-Path([string] $fromJson, [string] $fallback) {
  if ([string]::IsNullOrEmpty($fromJson)) { return $fallback }
  if ([IO.Path]::IsPathRooted($fromJson)) { return $fromJson }
  return (Join-Path $RdvHere $fromJson)
}
$RdvJsonData = $null; $RdvJsonLedger = $null; $RdvJsonLog = $null
if ($null -ne $RdvJson -and $null -ne $RdvJson.paths) {
  $RdvJsonData = $RdvJson.paths.dataDir
  $RdvJsonLedger = $RdvJson.paths.ledger
  $RdvJsonLog = $RdvJson.paths.log
}

$RdvData = Rdv-DataDir $RdvTokens $RdvHere (Rdv-Path $RdvJsonData $null)

$RdvLedger = Rdv-Opt $RdvTokens '-ledger'
if ([string]::IsNullOrEmpty($RdvLedger)) {
  $RdvLedger = Rdv-Path $RdvJsonLedger (Join-Path $RdvHere 'ReaderDataViewer-Ledger.xlsx')
}
$RdvLog = Rdv-Opt $RdvTokens '-log'
if ([string]::IsNullOrEmpty($RdvLog)) {
  $RdvLog = Rdv-Path $RdvJsonLog (Join-Path $RdvHere 'ReaderDataViewer.log')
}

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
$RdvRc = [Rdv3Program]::Run($RdvData, $RdvLedger, $RdvLog, $RdvConfig, $RdvCompileMs)
exit $RdvRc
