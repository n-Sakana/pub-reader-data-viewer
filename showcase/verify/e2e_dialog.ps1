# e2e_dialog.ps1 -- VERIFICATION ONLY. Presses a named button inside the Excel
# processes this run owns. Used for the two dialogs a normal open can show:
# the macro security bar, and Excel's "the document caused a serious error last
# time" prompt (which appears because a probe run had to be killed).
# It only ever looks inside the given process ids.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][int[]] $Pids,
    [Parameter(Mandatory=$true)][string] $ButtonPattern,
    [string] $WindowTextPattern = ''
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

$root = [System.Windows.Automation.AutomationElement]::RootElement
$desc = [System.Windows.Automation.TreeScope]::Descendants
$found = $false
foreach ($id in $Pids) {
    $cond = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ProcessIdProperty, $id)
    $wins = $root.FindAll([System.Windows.Automation.TreeScope]::Children, $cond)
    foreach ($win in $wins) {
        if ($WindowTextPattern) {
            $txtCond = New-Object System.Windows.Automation.PropertyCondition(
                [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
                [System.Windows.Automation.ControlType]::Text)
            $texts = $win.FindAll($desc, $txtCond)
            $hit = $false
            foreach ($t in $texts) { if ($t.Current.Name -match $WindowTextPattern) { $hit = $true; break } }
            if (-not $hit) { continue }
        }
        $bc = New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
            [System.Windows.Automation.ControlType]::Button)
        foreach ($b in $win.FindAll($desc, $bc)) {
            if ($b.Current.Name -match $ButtonPattern) {
                Write-Host "pressing [$($b.Current.Name)] in window [$($win.Current.Name)]"
                $b.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()
                $found = $true
                break
            }
        }
        if ($found) { break }
    }
    if ($found) { break }
}
if (-not $found) { Write-Host "no button matching /$ButtonPattern/ found" }
