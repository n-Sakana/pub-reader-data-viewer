param(
    [Parameter(Mandatory=$true)][long]$WindowHandle,
    [Parameter(Mandatory=$true)][int]$TargetProcessId,
    [Parameter(Mandatory=$true)][string]$Output,
    [int]$Count=12
)
$ErrorActionPreference='Stop'
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
$main=[System.Windows.Automation.AutomationElement]::FromHandle([IntPtr]$WindowHandle)
if ($main.Current.ProcessId -ne $TargetProcessId) { throw 'Window does not belong to the specified test process' }
$trace=New-Object 'System.Collections.Generic.List[object]'
function Node([string]$property,[string]$value) {
    $key=if ($property -eq 'id') {[System.Windows.Automation.AutomationElement]::AutomationIdProperty} else {[System.Windows.Automation.AutomationElement]::NameProperty}
    return $main.FindFirst([System.Windows.Automation.TreeScope]::Descendants,(New-Object System.Windows.Automation.PropertyCondition($key,$value)))
}
function InvokeNode($node) {
    if ($null -eq $node -or -not $node.Current.IsEnabled) { throw 'Required control is absent or disabled' }
    ([System.Windows.Automation.InvokePattern]$node.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)).Invoke()
}
function WaitUntil([scriptblock]$check,[string]$failure) {
    $deadline=(Get-Date).AddSeconds(10)
    do {
        if (& $check) { return }
        Start-Sleep -Milliseconds 100
    } while ((Get-Date) -lt $deadline)
    throw $failure
}
function Dialogs {
    return $main.FindAll([System.Windows.Automation.TreeScope]::Descendants,(New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ControlTypeProperty,[System.Windows.Automation.ControlType]::Window)))
}
try {
    for ($i=0; $i -lt $Count; $i++) {
        if (-not $main.Current.IsEnabled -or (Dialogs).Count -ne 0) { throw 'Start with an idle, enabled test window' }
        InvokeNode (Node 'id' 'b-set')
        WaitUntil { $null -ne (Node 'id' 'b-pick') } 'Settings did not open'
        InvokeNode (Node 'id' 'b-pick')
        WaitUntil { $null -ne (Node 'name' '閉じる') } 'Picker did not open'
        $mode=if ($i % 2 -eq 0) {'native-close'} else {'picker-close'}
        if ($mode -eq 'native-close') {
            $native=@((Dialogs) | Where-Object { $_.Current.NativeWindowHandle -ne 0 })
            if ($native.Count -ne 1) { throw 'Expected one owned native dialog' }
            ([System.Windows.Automation.WindowPattern]$native[0].GetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern)).Close()
        } else {
            InvokeNode (Node 'name' '閉じる')
            WaitUntil { $null -ne (Node 'id' 'b-pick') } 'Picker cancel did not return to settings'
            if ($main.Current.IsEnabled) { throw 'Main window was enabled while settings remained open' }
            InvokeNode (Node 'name' 'キャンセル')
        }
        WaitUntil { $main.Current.IsEnabled -and (Dialogs).Count -eq 0 } 'Closing the picker left the main window disabled or a dialog open'
        $searchBox=Node 'id' 'input'
        ([System.Windows.Automation.ValuePattern]$searchBox.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern)).SetValue('00000001')
        InvokeNode (Node 'id' 'b-clear')
        WaitUntil { ([System.Windows.Automation.ValuePattern]$searchBox.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern)).Current.Value -eq '' } 'Clear did not work after closing the picker'
        $trace.Add([pscustomobject]@{time=(Get-Date -Format o);cycle=$i;mode=$mode;mainEnabled=$main.Current.IsEnabled;dialogs=(Dialogs).Count;clearWorked=$true})
        Write-Output ('PASS picker cycle '+$i+' '+$mode)
    }
    Write-Output ('TOTAL '+$Count+' passed; 0 failed')
} finally {
    [IO.File]::WriteAllText([IO.Path]::GetFullPath($Output),($trace | ConvertTo-Json -Depth 4),(New-Object Text.UTF8Encoding($false)))
}
