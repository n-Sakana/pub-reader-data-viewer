# Capture the real WinForms control tree without showing a window.
[CmdletBinding()]
param(
  [string] $Root = "",
  [string] $OutPath = "",
  [switch] $Loaded,
  [switch] $Ng,
  [switch] $Waiting,
  [int] $PendingCount = 0,
  [switch] $DumpGeometry,
  [switch] $DumpButtons,
  [switch] $DumpStatus,
  [switch] $DumpJudgment
)
$ErrorActionPreference = 'Stop'
if ([string]::IsNullOrEmpty($Root)) { $Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path) }
if ([string]::IsNullOrEmpty($OutPath)) { $OutPath = Join-Path $Root 'work\ui-v3\app-headless.png' }

Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class RdvHeadlessDpi {
  [DllImport("user32.dll", SetLastError=true)] static extern bool SetProcessDpiAwarenessContext(IntPtr value);
  [DllImport("user32.dll")] public static extern uint GetDpiForWindow(IntPtr h);
  public static bool UseSystemAware() { return SetProcessDpiAwarenessContext(new IntPtr(-2)); }
  public static int LastError() { return Marshal.GetLastWin32Error(); }
}
"@
$dpiAwareSet = [RdvHeadlessDpi]::UseSystemAware()
$dpiAwareError = [RdvHeadlessDpi]::LastError()

. (Join-Path $Root 'build\sources.ps1')
$usings = New-Object System.Collections.Specialized.OrderedDictionary
$bodies = New-Object System.Text.StringBuilder
foreach ($source in $RdvSources) {
  $text = [IO.File]::ReadAllText((Join-Path $Root "src\csharp\$source"), [Text.Encoding]::UTF8)
  foreach ($line in ($text -split "`r?`n")) {
    if ($line -match '^\s*using\s+[A-Za-z_][A-Za-z0-9_.]*\s*;\s*$') {
      $key = $line.Trim()
      if (-not $usings.Contains($key)) { $usings.Add($key, $true) }
    } else { [void]$bodies.AppendLine($line) }
  }
}
$code = (($usings.Keys | ForEach-Object { $_ }) -join "`r`n") + "`r`n`r`n" + $bodies.ToString()
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
  [System.Xml.XmlReader].Assembly.Location
)
Add-Type -TypeDefinition $code -ReferencedAssemblies $refs -Language CSharp

$configPath = Join-Path $Root 'src\config\settings.json'
$dataDir = Join-Path $Root 'dist\app-csharp\data'
$cfg = [Rdv3Config]::Load($configPath)
$heads = New-Object 'string[][]' $cfg.Data.Tables.Count
for ($i = 0; $i -lt $cfg.Data.Tables.Count; $i++) {
  $heads[$i] = [Rdv3Table]::ReadHead((Join-Path $dataDir $cfg.Data.Tables[$i].File), $cfg.Data.Enc)
}
$cfg.Data.Bind($heads)
$form = New-Object Rdv3Form $cfg.Screen
$form.SetFields((New-Object Rdv3Fields (,$cfg.Data.ColumnRefs)))
$targetName = $cfg.Targets[0].Name
if ($Waiting) {
  $form.SetState([Rdv3Text]::StateWaitingFmt.Replace('{name}', $targetName))
} else {
  $form.SetState([Rdv3Text]::StateReady)
}
$form.SetWatch($targetName, [Rdv3Text]::WatchConnectedFmt.Replace('{title}', $targetName))
$form.SetLedger([Rdv3Text]::LedgerSegFmt.Replace('{file}', 'ReaderDataViewer-Ledger.xlsx').Replace('{n}', '100,000'), '100,000', '09-03 11:42')
$form.SetPendingCount($PendingCount)
if ($Loaded -or $Ng) {
  $merged = [Rdv3Ledger]::BuildFromCsv($cfg.Data, $dataDir)
  $line = $merged.Lines[0]
  if ($Ng) {
    $statusColumn = $cfg.Data.IndexOf('B.b_status')
    foreach ($candidateLine in $merged.Lines) {
      if ([Rdv3Ledger]::FieldOf($candidateLine, $statusColumn) -eq 'HOLD') { $line = $candidateLine; break }
    }
  }
  $candidate = New-Object Rdv3CandRow
  $candidate.Line = $line
  $candidate.Stored = $cfg.Screen.Work.InitialStored
  $list = New-Object 'System.Collections.Generic.List[Rdv3CandRow]'
  $list.Add($candidate)
  $key = [Rdv3Ledger]::FieldOf($candidate.Line, $cfg.Data.SearchCols[0])
  $form.SetKeyText($key)
  $form.ShowCandidates($key, $list, 1)
  $form.SelectCandidate(0)
}
function PrepareControl([System.Windows.Forms.Control] $control) {
  $control.CreateControl()
  foreach ($child in $control.Controls) { PrepareControl $child }
  $control.PerformLayout()
}
PrepareControl $form
$form.PerformLayout()
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
$form.Location = New-Object Drawing.Point(-32000, -32000)
$form.ShowInTaskbar = $false
$form.Show()
[System.Windows.Forms.Application]::DoEvents()
if ($DumpGeometry) { Write-Output $form.GeometryDump() }
if ($DumpButtons) {
  Write-Output ("form client={0}x{1} dpi={2} setAware={3} error={4} autoScale={5} currentScale={6}" -f $form.ClientSize.Width, $form.ClientSize.Height, [RdvHeadlessDpi]::GetDpiForWindow($form.Handle), $dpiAwareSet, $dpiAwareError, $form.AutoScaleDimensions, $form.CurrentAutoScaleDimensions)
  $todo = New-Object 'System.Collections.Generic.Stack[System.Windows.Forms.Control]'
  $badButtons = New-Object 'System.Collections.Generic.List[string]'
  $searchFigure = $null
  $searchInputRow = $null
  $todo.Push($form)
  while ($todo.Count -gt 0) {
    $at = $todo.Pop()
    foreach ($child in $at.Controls) {
      $todo.Push($child)
      if ($child.Name -eq 'section0.figure') { $searchFigure = $child }
      if ($child.Name -eq 'section0.inputRow') { $searchInputRow = $child }
      if ($child -is [System.Windows.Forms.ButtonBase] -and $child.Name -like 'button.*') {
        $preferred = $child.GetPreferredSize((New-Object Drawing.Size(10000, 10000)))
        Write-Output ("button {0} actual={1}x{2} preferred={3}x{4} autosize={5} text={6}" -f $child.Name, $child.Width, $child.Height, $preferred.Width, $preferred.Height, $child.AutoSize, $child.Text)
        if (-not $child.AutoSize -or $child.Width -lt $preferred.Width -or $child.Height -lt $preferred.Height) {
          $badButtons.Add($child.Name)
        }
      }
    }
  }
  if ($null -ne $searchFigure -and $null -ne $searchInputRow) {
    $figureTable = $searchFigure.Parent
    $splitTable = $figureTable.Parent
    Write-Output ("search split={0}x{1} columns={2} figureTable={3} inputRow={4}" -f $splitTable.Width, $splitTable.Height, ($splitTable.GetColumnWidths() -join ','), $figureTable.Bounds, $searchInputRow.Bounds)
  }
  if ($badButtons.Count -gt 0) { throw ("buttons do not fit their preferred size: " + ($badButtons -join ', ')) }
}
if ($DumpStatus) {
  $statusBar = $form.Controls.Find('statusBar', $true)[0]
  $boldFont = New-Object System.Drawing.Font($form.Font, [System.Drawing.FontStyle]::Bold)
  try {
    for ($i = 0; $i -lt $statusBar.Items.Count; $i++) {
      $panel = $statusBar.Items[$i]
      $measureFont = if ($i -eq 0 -or ($i -eq 3 -and $PendingCount -ne 0)) { $boldFont } else { $form.Font }
      $needed = [System.Windows.Forms.TextRenderer]::MeasureText($panel.Text, $measureFont).Width
      Write-Output ("status[{0}] text={1} actual={2} needed={3} spring={4}" -f $i, $panel.Text, $panel.Width, $needed, $panel.Spring)
      if ($i -ne 2 -and $panel.Width -lt $needed) { throw ("status panel text is clipped: " + $i) }
    }
  } finally { $boldFont.Dispose() }
}
if ($DumpJudgment) {
  $judgment = $form.Controls.Find('section4.judgment', $true)[0]
  $color = $judgment.ForeColor
  Write-Output ("judgment text={0} font={1}pt/{2} color=#{3:X2}{4:X2}{5:X2} bounds={6}" -f $judgment.Text, $judgment.Font.SizeInPoints, $judgment.Font.Style, $color.R, $color.G, $color.B, $judgment.Bounds)
  if ($judgment.Text -eq [Rdv3Text]::Unsearched) {
    if ([Math]::Abs($judgment.Font.SizeInPoints - 12.0) -gt 0.01 -or $color.R -ne 96 -or $color.G -ne 96 -or $color.B -ne 96) { throw 'unsearched judgment style does not match the mock' }
  } else {
    $expected = if ($judgment.Text -eq 'OK') { [System.Drawing.Color]::DarkGreen } else { [System.Drawing.Color]::Maroon }
    if ([Math]::Abs($judgment.Font.SizeInPoints - 15.0) -gt 0.01 -or $color.ToArgb() -ne $expected.ToArgb()) { throw 'result judgment style does not match the mock' }
  }
}
$bitmap = New-Object System.Drawing.Bitmap($form.Width, $form.Height)
$form.DrawToBitmap($bitmap, (New-Object Drawing.Rectangle(0, 0, $bitmap.Width, $bitmap.Height)))
$folder = Split-Path -Parent $OutPath
if (-not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Path $folder | Out-Null }
$bitmap.Save($OutPath, [Drawing.Imaging.ImageFormat]::Png)
$bitmap.Dispose()
$form.Hide()
$form.Dispose()
Write-Output $OutPath
