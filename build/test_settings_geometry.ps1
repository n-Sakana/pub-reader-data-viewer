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
function Invoke-DoubleClick([System.Windows.Forms.Control] $control) {
  $flags = [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::NonPublic
  $method = [System.Windows.Forms.Control].GetMethod('OnDoubleClick', $flags)
  [void]$method.Invoke($control, @([EventArgs]::Empty))
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
$keyFigure = Named $main 'section0.figure'
$fieldValue = Named $main 'section1.item0.value0'
$fieldLabel = Named $main 'section1.item0.label0'
$judgment = Named $main 'section4.judgment'
$judgmentSub = Named $main 'section4.sub'
$judgmentBand = Named $main 'section4'
Check 'record values are bold while their labels stay regular' ((($fieldValue.Font.Style -band [Drawing.FontStyle]::Bold) -ne 0) -and (($fieldLabel.Font.Style -band [Drawing.FontStyle]::Bold) -eq 0)) ("value=" + $fieldValue.Font.Style + " label=" + $fieldLabel.Font.Style)
Check 'the three emphasis sizes are taken from JSON' (([Math]::Abs($keyFigure.Font.Size - $cfg.Screen.KeyValueFontSize) -lt 0.1) -and ([Math]::Abs($judgment.Font.Size - $cfg.Screen.UnsearchedFontSize) -lt 0.1)) ("key=" + $keyFigure.Font.Size + " judgment=" + $judgment.Font.Size)
$judgmentCenter = $judgment.Left + ($judgment.Width / 2.0)
$bandCenter = $judgmentBand.ClientSize.Width / 2.0
Check 'the judgment itself is centered and its subtext follows right' (([Math]::Abs($judgmentCenter - $bandCenter) -le 1.0) -and ($judgmentSub.Left -ge $judgment.Right)) ("judgment=" + $judgmentCenter + " band=" + $bandCenter + " sub=" + $judgmentSub.Left)
Check 'empty main-screen fields stay blank without a dash substitute' ((@($main | Where-Object { [string]::Equals($_.Text, [char]0x2014) }).Count -eq 0) -and ($null -eq [Rdv3Text].GetField('Dash'))) 'dash substitute remains'
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
$framePadding = [Rdv3Metrics]::FramedPadding()
$fieldPadding = [Rdv3Metrics]::FieldListPadding()
$framedSections = @('section0', 'section2', 'section3') | ForEach-Object { Named $main $_ }
$fieldSections = @('section1.item0', 'section1.item1') | ForEach-Object { Named $main $_ }
$framePaddingOk = (@($framedSections | Where-Object { -not $_.Padding.Equals($framePadding) }).Count -eq 0) -and
  (@($fieldSections | Where-Object { -not $_.Padding.Equals($fieldPadding) }).Count -eq 0)
Check 'the card gap supplies every framed-section inset' $framePaddingOk ("frame=" + $framePadding + " fields=" + $fieldPadding)
$definedGap = [int][Math]::Round($cfg.Screen.Gap)
$mainSections = @('section0', 'section1', 'section2', 'section3', 'section4') | ForEach-Object { Named $main $_ }
$sectionGaps = for ($i = 1; $i -lt $mainSections.Count; $i++) { $mainSections[$i].Top - $mainSections[$i - 1].Bottom }
$commandPanel = Named $main 'commandBar'
$commandButtons = Named $main 'commandButtons'
$sendPanel = Named $main 'sendBar'
$sendSeparator = Named $main 'sendBar.separator'
$aboveActions = $commandPanel.Top + $commandButtons.Top - $mainSections[-1].Bottom
$belowActions = $sendPanel.Top + $sendSeparator.Top - ($commandPanel.Top + $commandButtons.Bottom)
$allMainGaps = @($sectionGaps) + @($aboveActions, $belowActions)
Check 'all main section and bottom-action gaps use the card gap' ((@($allMainGaps | Where-Object { $_ -ne $definedGap }).Count) -eq 0) ("expected " + $definedGap + " actual " + ($allMainGaps -join ','))
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

$export = [Rdv3ExportForm]::ForCheck($cfg.Data, $cfg.Screen, (Join-Path $Root 'dist\app-csharp'))
Prepare $export
$ec = @(Controls-Of $export)
Check 'export uses two standard ListBox controls' ((@($ec | Where-Object { $_ -is [System.Windows.Forms.ListBox] }).Count) -eq 2) 'list boxes'
Check 'export has both move buttons and reset beside the heading' (($null -ne (Named $ec 'export.right')) -and ($null -ne (Named $ec 'export.left')) -and ($null -ne (Named $ec 'export.reset'))) 'export actions'
Check 'export destination is an editable white TextBox' (((Named $ec 'export.destination') -is [System.Windows.Forms.TextBox]) -and (-not (Named $ec 'export.destination').ReadOnly) -and ((Named $ec 'export.destination').BackColor -eq [System.Drawing.SystemColors]::Window)) 'destination'
$filterGroup = Named $ec 'export.filterGroup'
$filterList = Named $ec 'export.filter.list'
Check 'export filter editor uses only standard input and list controls' (($filterGroup -is [System.Windows.Forms.GroupBox]) -and ($filterList -is [System.Windows.Forms.ListView]) -and ($filterList.View -eq [System.Windows.Forms.View]::Details) -and (@($ec | Where-Object { $_ -is [System.Windows.Forms.ComboBox] }).Count -eq 2) -and (@($ec | Where-Object { $_ -is [System.Windows.Forms.DateTimePicker] }).Count -eq 2)) 'filter controls'
Check 'export filter has add and remove actions' (((Named $ec 'export.filter.add') -is [System.Windows.Forms.Button]) -and ((Named $ec 'export.filter.remove') -is [System.Windows.Forms.Button])) 'filter actions'
$available = Named $ec 'export.available'
$selected = Named $ec 'export.selected'
$configuredDefaults = @($cfg.Screen.ExportDefaultFields)
$selectedRefs = @(foreach ($item in $selected.Items) { $item.Ref })
Check 'export starts with the five configured default fields in order' (($available.Items.Count -eq 16) -and (($selectedRefs -join '|') -eq ($configuredDefaults -join '|'))) (("counts " + $available.Items.Count + '/' + $selected.Items.Count) + ' refs=' + ($selectedRefs -join '|'))
$available.SelectedIndex = 0
Invoke-Click (Named $ec 'export.right')
Check 'move right transfers the selected output field' (($available.Items.Count -eq 15) -and ($selected.Items.Count -eq 6)) ("counts " + $available.Items.Count + '/' + $selected.Items.Count)
$selected.SelectedIndex = $selected.Items.Count - 1
Invoke-DoubleClick $selected
Check 'double-clicking a selected output field moves it back left' (($available.Items.Count -eq 16) -and ($selected.Items.Count -eq 5)) ("counts " + $available.Items.Count + '/' + $selected.Items.Count)
$available.SelectedIndex = $available.Items.Count - 1
Invoke-DoubleClick $available
Check 'double-clicking an available field moves it right' (($available.Items.Count -eq 15) -and ($selected.Items.Count -eq 6)) ("counts " + $available.Items.Count + '/' + $selected.Items.Count)
Invoke-Click (Named $ec 'export.reset')
$resetRefs = @(foreach ($item in $selected.Items) { $item.Ref })
Check 'reset restores the five configured default fields in order' (($available.Items.Count -eq 16) -and (($resetRefs -join '|') -eq ($configuredDefaults -join '|'))) (("counts " + $available.Items.Count + '/' + $selected.Items.Count) + ' refs=' + ($resetRefs -join '|'))
(Named $ec 'export.filter.value1').Text = 'SAMPLE'
Invoke-Click (Named $ec 'export.filter.add')
Check 'a text filter is added to the visible condition list' (($filterList.Items.Count -eq 1) -and ($filterList.Items[0].SubItems[1].Text -eq 'を含む') -and ($filterList.Items[0].SubItems[2].Text -eq 'SAMPLE')) 'filter was not listed'
$filterList.Items[0].Selected = $true
Invoke-Click (Named $ec 'export.filter.remove')
Check 'the selected filter can be removed' ($filterList.Items.Count -eq 0) ("filters " + $filterList.Items.Count)
$export.Dispose()

$typedExport = [Rdv3ExportForm]::ForCheck($cfg.Data, $cfg.Screen, (Join-Path $Root 'dist\app-csharp'))
$typedExport.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
$typedExport.Location = New-Object Drawing.Point(-32000, -32000)
$typedExport.ShowInTaskbar = $false
Prepare $typedExport
$typedExport.Show()
[System.Windows.Forms.Application]::DoEvents()
$tec = @(Controls-Of $typedExport)
$typedField = Named $tec 'export.filter.field'
for ($i = 0; $i -lt $typedField.Items.Count; $i++) { if ($typedField.Items[$i].Ref -eq 'B.b_date') { $typedField.SelectedIndex = $i; break } }
[System.Windows.Forms.Application]::DoEvents()
$dateInputs = @($tec | Where-Object { $_ -is [System.Windows.Forms.DateTimePicker] -and $_.Visible })
$dateOperator = Named $tec 'export.filter.operator'
Check 'a shipped date filter shows range and two calendar inputs' (($dateOperator.Items.Count -eq 1) -and ($dateOperator.Items[0].Code -eq 'range') -and ($dateInputs.Count -eq 2) -and ($dateInputs[0].CustomFormat -eq 'yyyyMMdd') -and ((Named $tec 'export.filter.value2').Visible -eq $false)) ("date inputs " + $dateInputs.Count)
for ($i = 0; $i -lt $typedField.Items.Count; $i++) { if ($typedField.Items[$i].Ref -eq 'B.b_qty') { $typedField.SelectedIndex = $i; break } }
[System.Windows.Forms.Application]::DoEvents()
$numberInputs = @($tec | Where-Object { $_ -is [System.Windows.Forms.TextBox] -and $_.Name -match '^export\.filter\.value[12]$' -and $_.Visible })
Check 'a shipped number filter shows range and two numeric text inputs' (($dateOperator.Items.Count -eq 1) -and ($dateOperator.Items[0].Code -eq 'range') -and ($numberInputs.Count -eq 2) -and (@($tec | Where-Object { $_ -is [System.Windows.Forms.DateTimePicker] -and $_.Visible }).Count -eq 0)) ("number inputs " + $numberInputs.Count)
$typedExport.Hide()
$typedExport.Dispose()

$confirm = [Rdv3ConfirmForm]::ForCheck([Rdv3Text]::ConfirmUpdateTitle, [Rdv3Text]::UpdateConfirmBody($cfg.Data.WorkStateOnSourceChange, $cfg.Screen.Work.InitialState.Text))
Prepare $confirm
$cf = @(Controls-Of $confirm)
Check 'confirm uses native Label and Button controls' ((@($cf | Where-Object { $_ -is [System.Windows.Forms.Label] }).Count -ge 1) -and (@($cf | Where-Object { $_ -is [System.Windows.Forms.Button] }).Count -eq 2)) 'confirm controls'
$confirmBody = Named $cf 'confirm.body'
$confirmButtons = Named $cf 'confirm.buttons'
Check 'confirm height follows its text instead of leaving the old empty area' (($confirm.ClientSize.Height -lt 190) -and ($confirmBody.Height + $confirmButtons.Height -eq $confirm.ClientSize.Height)) ("client=" + $confirm.ClientSize.Height + " parts=" + ($confirmBody.Height + $confirmButtons.Height))
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
