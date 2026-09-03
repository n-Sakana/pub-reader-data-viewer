# Verify the native-control UI contract without putting a window on screen.
[CmdletBinding()]
param([string] $Root = "")
$ErrorActionPreference = 'Stop'
if ([string]::IsNullOrEmpty($Root)) { $Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path) }

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
Write-Output 'compile ok'

$script:pass = 0
$script:fail = 0
function Check([string] $name, [bool] $condition, [string] $detail) {
  if ($condition) { $script:pass++; Write-Output ("  ok   " + $name) }
  else { $script:fail++; Write-Output ("  FAIL " + $name + " (" + $detail + ")") }
}
function Controls-Of($root) {
  $result = New-Object System.Collections.ArrayList
  foreach ($control in $root.Controls) {
    [void]$result.Add($control)
    foreach ($child in (Controls-Of $control)) { [void]$result.Add($child) }
  }
  return $result
}
function Prepare($control) {
  $control.CreateControl()
  foreach ($child in $control.Controls) { Prepare $child }
  $control.PerformLayout()
}
function Named($controls, [string] $name) {
  return @($controls | Where-Object { $_.Name -eq $name })[0]
}
function Invoke-Click([System.Windows.Forms.Button] $button) {
  # PerformClick ignores hidden forms. Invoke the protected native-control
  # click path directly so the verification remains entirely headless.
  $flags = [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::NonPublic
  $method = $button.GetType().GetMethod('OnClick', $flags)
  [void]$method.Invoke($button, @([EventArgs]::Empty))
}

$uiFiles = @('Rdv3Ui.cs', 'Rdv3Modals.cs', 'Rdv3Settings.cs')
$uiText = ($uiFiles | ForEach-Object { [IO.File]::ReadAllText((Join-Path $Root "src\csharp\$_"), [Text.Encoding]::UTF8) }) -join "`n"
$appText = [IO.File]::ReadAllText((Join-Path $Root 'src\csharp\Rdv3App.cs'), [Text.Encoding]::UTF8)
Check 'the three UI files contain no OnPaint override' (-not ($uiText -match '\bOnPaint\s*\(')) 'custom paint remains'
Check 'the three UI files contain no Graphics drawing' (-not ($uiText -match '\bGraphics\b|DrawRectangle|FillRectangle|DrawString')) 'drawing API remains'
Check 'visual styles are not enabled' (-not ($appText -match 'EnableVisualStyles|SetCompatibleTextRenderingDefault')) 'modern rendering call remains'

$cfg = [Rdv3Config]::Load((Join-Path $Root 'src\config\settings.json'))
$form = New-Object Rdv3Form $cfg.Screen
Prepare $form
$main = @(Controls-Of $form)
$groups = @($main | Where-Object { $_ -is [System.Windows.Forms.GroupBox] })
$readOnly = @($main | Where-Object { $_ -is [System.Windows.Forms.TextBox] -and $_.ReadOnly })
$shallowValues = @($main | Where-Object {
  $_ -is [System.Windows.Forms.Label] -and $_.Name -match '\.(figure|value[0-9]+)$' -and
  $_.BorderStyle -eq [System.Windows.Forms.BorderStyle]::Fixed3D -and
  $_.BackColor -eq [System.Drawing.SystemColors]::Control
})
$editable = @($main | Where-Object { $_ -is [System.Windows.Forms.TextBox] -and -not $_.ReadOnly })
$status = @($main | Where-Object { $_ -is [System.Windows.Forms.StatusBar] })
$strips = @($main | Where-Object { $_ -is [System.Windows.Forms.StatusStrip] })
$workButton = Named $main 'button.workState'
$sendButton = Named $main 'button.sendChanges'
$sendLabel = Named $main 'sendBar.pending'
Check 'main screen is built from GroupBox controls' ($groups.Count -ge 5) ("count " + $groups.Count)
Check 'main screen has one editable white search box' (($editable.Count -eq 1) -and ($editable[0].BackColor -eq [System.Drawing.SystemColors]::Window)) ("count " + $editable.Count)
Check 'single-line record values use shallow Fixed3D labels' ($shallowValues.Count -ge 11) ("count " + $shallowValues.Count)
Check 'long record values keep scrollable read-only text boxes' (($readOnly.Count -eq 2) -and (@($readOnly | Where-Object { -not $_.Multiline -or $_.ScrollBars -ne [System.Windows.Forms.ScrollBars]::Vertical -or $_.BackColor -ne [System.Drawing.SystemColors]::Control }).Count -eq 0)) ("count " + $readOnly.Count)
Check 'work state is a CheckBox with button appearance' (($null -ne $workButton) -and ($workButton -is [System.Windows.Forms.CheckBox]) -and ($workButton.Appearance -eq [System.Windows.Forms.Appearance]::Button)) 'wrong control'
Check 'the status strip has exactly four segments' (($strips.Count -eq 1) -and ($strips[0].Items.Count -eq 4)) 'status segments'
Check 'the third segment is the one that springs' (($strips.Count -eq 1) -and ($strips[0].Items[2].Spring) -and (@(0, 1, 3) | Where-Object { $strips[0].Items[$_].Spring }).Count -eq 0) 'wrong spring'
Check 'every segment carries a sunken border' (($strips.Count -eq 1) -and (@($strips[0].Items | Where-Object { $_.BorderSides -ne [System.Windows.Forms.ToolStripStatusLabelBorderSides]::All }).Count -eq 0)) 'segment border'
Check 'the retired StatusBar control is gone' ($status.Count -eq 0) 'old status bar remains'
Check 'there is no menu strip' ((@($main | Where-Object { $_ -is [System.Windows.Forms.MenuStrip] }).Count) -eq 0) 'menu strip found'
$actions = @('button.tableExport', 'button.updateRecords', 'button.deleteRecords', 'button.settings')
Check 'all four bottom actions exist' ((@($actions | Where-Object { $null -ne (Named $main $_) }).Count) -eq 4) 'button missing'
$sendLabelOk = $sendLabel -is [System.Windows.Forms.Label]
$sendButtonOk = $sendButton -is [System.Windows.Forms.Button]
$sendTextOk = [string]::Equals($sendButton.Text, '送信(&U)')
Check 'the separate send band has its count and send button' ($sendLabelOk -and $sendButtonOk -and $sendTextOk) ("label=" + $sendLabelOk + " button=" + $sendButtonOk + " text=" + $sendTextOk)
# A pinned number went stale every time a font or a band changed. What the
# start height has to earn is that the whole screen is visible without the
# scroll bar, so assert that instead of the number.
$contentHost = Named $main 'contentHost'
$contentPanel = Named $main 'content'
$heightDeclared = [int][Math]::Round($cfg.Screen.StartHeight)
Check 'the window starts at the height the settings declare' ($form.ClientSize.Height -eq $heightDeclared) ("height " + $form.ClientSize.Height + " declared " + $heightDeclared)
Check 'the whole screen fits at the start height, unscrolled' ($contentPanel.Height -le $contentHost.ClientSize.Height) ("content " + $contentPanel.Height + " host " + $contentHost.ClientSize.Height)
$form.SetPendingCount(2)
$pendingTextOk = [string]::Equals($sendLabel.Text, '未送信 2 件')
$pendingColorOk = $sendLabel.ForeColor.ToArgb() -eq [Drawing.Color]::Maroon.ToArgb()
$pendingBoldOk = ($sendLabel.Font.Style -band [Drawing.FontStyle]::Bold) -ne 0
Check 'a nonzero unsent count is maroon and bold' ($pendingTextOk -and $pendingColorOk -and $pendingBoldOk) ("text=" + $pendingTextOk + " color=" + $pendingColorOk + " bold=" + $pendingBoldOk)
$form.Dispose()

$settings = [Rdv3SettingsForm]::ForCheck($cfg.Clone())
Prepare $settings
$sc = @(Controls-Of $settings)
Check 'settings has the three mock groups' ((@($sc | Where-Object { $_ -is [System.Windows.Forms.GroupBox] }).Count) -eq 3) 'group count'
$namedEdits = @('settings.path0', 'settings.path1', 'settings.path2', 'settings.pattern') | ForEach-Object { Named $sc $_ }
Check 'settings paths and search fields are editable white controls' ((@($namedEdits | Where-Object { $_ -is [System.Windows.Forms.TextBox] -and -not $_.ReadOnly -and $_.BackColor -eq [System.Drawing.SystemColors]::Window }).Count) -eq 4) 'editable fields'
Check 'settings target and read fields are read-only gray controls' ((@($sc | Where-Object { $_ -is [System.Windows.Forms.TextBox] -and $_.ReadOnly -and $_.BackColor -eq [System.Drawing.SystemColors]::Control }).Count) -eq 2) 'read-only fields'
$settings.Dispose()

$candidateRows = New-Object 'System.Collections.Generic.List[Rdv3CandRow]'
$candidate = New-Object Rdv3CandRow
$candidate.Line = ((1..28 | ForEach-Object { '' }) -join "`t")
$candidate.Stored = $cfg.Screen.Work.InitialStored
$candidateRows.Add($candidate)
$candidates = [Rdv3CandidatesForm]::ForCheck($cfg.Screen.Candidates, (New-Object Rdv3Fields (,$cfg.Data.ColumnRefs)), $cfg.Screen.Work, $candidateRows, 1, 0)
Prepare $candidates
$cc = @(Controls-Of $candidates)
$candidateList = Named $cc 'candidates.list'
Check 'candidate records use a Details ListView' (($candidateList -is [System.Windows.Forms.ListView]) -and ($candidateList.View -eq [System.Windows.Forms.View]::Details)) 'candidate list'
Check 'candidate ListView columns come from settings' ($candidateList.Columns.Count -eq $cfg.Screen.Candidates.Columns.Count) 'column count'
$candidates.Dispose()

$dataDir = Join-Path $Root 'dist\app-csharp\data'
$ledger = Join-Path $Root 'dist\app-csharp\ReaderDataViewer-Ledger.xlsx'
$process = [Rdv3ProcessForm]::ForCheck($cfg.Data, 'merge-ledger', $dataDir, $ledger)
Prepare $process
$pc = @(Controls-Of $process)
$processGroups = @($pc | Where-Object { $_ -is [System.Windows.Forms.GroupBox] })
$processLists = @($pc | Where-Object { $_ -is [System.Windows.Forms.ListView] })
Check 'process dialog has exactly three groups and no job picker' (($processGroups.Count -eq 3) -and (@($pc | Where-Object { $_.Name -match 'job' }).Count -eq 0)) ("groups " + $processGroups.Count)
Check 'inputs and steps use the same Details ListView type' (($processLists.Count -eq 2) -and (@($processLists | Where-Object { $_.View -ne [System.Windows.Forms.View]::Details }).Count -eq 0)) 'list types'
$stepList = Named $pc 'process.steps'
Check 'step table has the seven defined columns' ($stepList.Columns.Count -eq 7) 'step columns'
Check 'step text is generated from operation, keys, condition, and output labels' (($stepList.Items[0].SubItems[1].Text -eq '結合') -and ($stepList.Items[0].SubItems[4].Text -eq '番号1 = 番号1') -and ($stepList.Items[0].SubItems[5].Text -eq '左外部')) (($stepList.Items[0].SubItems | ForEach-Object { $_.Text }) -join '|')
Check 'output contains three read-only rows' ((@($pc | Where-Object { $_.Name -match '^process\.outputValue' -and $_.ReadOnly }).Count) -eq 3) 'output rows'
$process.Dispose()

$export = [Rdv3ExportForm]::ForCheck($cfg.Data, (Join-Path $Root 'dist\app-csharp'))
Prepare $export
$ec = @(Controls-Of $export)
Check 'export uses two standard ListBox controls' ((@($ec | Where-Object { $_ -is [System.Windows.Forms.ListBox] }).Count) -eq 2) 'list boxes'
Check 'export has both move buttons and reset beside the heading' (($null -ne (Named $ec 'export.right')) -and ($null -ne (Named $ec 'export.left')) -and ($null -ne (Named $ec 'export.reset'))) 'export actions'
Check 'export destination is an editable white TextBox' (((Named $ec 'export.destination') -is [System.Windows.Forms.TextBox]) -and (-not (Named $ec 'export.destination').ReadOnly) -and ((Named $ec 'export.destination').BackColor -eq [System.Drawing.SystemColors]::Window)) 'destination'
$available = Named $ec 'export.available'
$selected = Named $ec 'export.selected'
Check 'export starts with the mock 21 fields and five defaults' (($available.Items.Count -eq 16) -and ($selected.Items.Count -eq 5)) ("counts " + $available.Items.Count + '/' + $selected.Items.Count)
$available.SelectedIndex = 0
Invoke-Click (Named $ec 'export.right')
Check 'move right transfers the selected output field' (($available.Items.Count -eq 15) -and ($selected.Items.Count -eq 6)) ("counts " + $available.Items.Count + '/' + $selected.Items.Count)
Invoke-Click (Named $ec 'export.reset')
Check 'reset restores the five default output fields' (($available.Items.Count -eq 16) -and ($selected.Items.Count -eq 5)) ("counts " + $available.Items.Count + '/' + $selected.Items.Count)
$export.Dispose()

$confirm = [Rdv3ConfirmForm]::ForCheck([Rdv3Text]::ConfirmUpdateTitle, [Rdv3Text]::UpdateConfirmBody($cfg.Data.WorkStateOnSourceChange, $cfg.Screen.Work.InitialState.Text))
Prepare $confirm
$cf = @(Controls-Of $confirm)
Check 'confirm uses native Label and Button controls' ((@($cf | Where-Object { $_ -is [System.Windows.Forms.Label] }).Count -ge 1) -and (@($cf | Where-Object { $_ -is [System.Windows.Forms.Button] }).Count -eq 2)) 'confirm controls'
$confirm.Dispose()

$picker = [Rdv3PickerForm]::ForCheck()
Prepare $picker
$pk = @(Controls-Of $picker)
Check 'picker displays six read-only standard fields' ((@($pk | Where-Object { $_ -is [System.Windows.Forms.TextBox] -and $_.ReadOnly }).Count) -eq 6) 'picker fields'
$picker.Dispose()

Write-Output ''
Write-Output ("{0} passed, {1} failed" -f $script:pass, $script:fail)
if ($script:fail -gt 0) { Write-Output 'RESULT: FAIL'; exit 1 }
Write-Output 'RESULT: PASS'
