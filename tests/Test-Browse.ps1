param(
    [Parameter(Mandatory=$true)][long]$WindowHandle,
    [Parameter(Mandatory=$true)][int]$TargetProcessId,
    [Parameter(Mandatory=$true)][string]$Output
)
# Real-window check of the settings dialog's browse buttons: each native picker
# must be owned by the settings surface (not unowned, not owned by the main
# window), the surface must be disabled while the picker is open, and the
# surface must answer again once the picker is closed. Run against a test
# window only; never against an operator's window.
$ErrorActionPreference='Stop'
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -TypeDefinition @'
using System;
using System.Text;
using System.Runtime.InteropServices;
public static class BrowseProbe {
  [DllImport("user32.dll")] public static extern IntPtr GetWindow(IntPtr h, uint cmd);
  [DllImport("user32.dll")] public static extern bool IsWindowEnabled(IntPtr h);
  [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  public static string ClassOf(IntPtr h) { StringBuilder sb = new StringBuilder(256); GetClassName(h, sb, 256); return sb.ToString(); }
}
'@
$main=[System.Windows.Automation.AutomationElement]::FromHandle([IntPtr]$WindowHandle)
if ($main.Current.ProcessId -ne $TargetProcessId) { throw 'Window does not belong to the specified test process' }
$root=[System.Windows.Automation.AutomationElement]::RootElement
$trace=New-Object 'System.Collections.Generic.List[object]'
# Every window of the process UIA can see: unowned ones are children of the
# desktop, owned ones (the dialog surface, an owned picker) sit under their owner.
function TopLevel {
    $cond=New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ProcessIdProperty,$TargetProcessId)
    $found=@($root.FindAll([System.Windows.Automation.TreeScope]::Children,$cond))
    $windowType=New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ControlTypeProperty,[System.Windows.Automation.ControlType]::Window)
    foreach ($w in $main.FindAll([System.Windows.Automation.TreeScope]::Descendants,$windowType)) {
        if ($w.Current.NativeWindowHandle -ne 0 -and ($found | Where-Object { $_.Current.NativeWindowHandle -eq $w.Current.NativeWindowHandle }).Count -eq 0) { $found += $w }
    }
    return $found
}
function Node($scope,[string]$property,[string]$value) {
    $key=if ($property -eq 'id') {[System.Windows.Automation.AutomationElement]::AutomationIdProperty} else {[System.Windows.Automation.AutomationElement]::NameProperty}
    return $scope.FindFirst([System.Windows.Automation.TreeScope]::Descendants,(New-Object System.Windows.Automation.PropertyCondition($key,$value)))
}
function InvokeNode($node) {
    if ($null -eq $node -or -not $node.Current.IsEnabled) { throw 'Required control is absent or disabled' }
    try { ([System.Windows.Automation.InvokePattern]$node.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)).Invoke() }
    catch [System.Windows.Automation.ElementNotAvailableException] { }
}
function WaitUntil([scriptblock]$check,[string]$failure,[int]$seconds=10) {
    $deadline=(Get-Date).AddSeconds($seconds)
    do { if (& $check) { return }; Start-Sleep -Milliseconds 100 } while ((Get-Date) -lt $deadline)
    throw $failure
}
function SettingsWindow {
    foreach ($w in TopLevel) { if ($w.Current.Name -eq '設定' -and $w.Current.NativeWindowHandle -ne 0) { return $w } }
    return $null
}
try {
    if (-not $main.Current.IsEnabled) { throw 'Start with an idle, enabled test window' }
    InvokeNode (Node $main 'id' 'b-set')
    WaitUntil { $null -ne (SettingsWindow) } 'Settings did not open'
    $settings=SettingsWindow
    $settingsHwnd=[IntPtr]$settings.Current.NativeWindowHandle
    $known=@(TopLevel | ForEach-Object { $_.Current.NativeWindowHandle })
    $buttons=@($settings.FindAll([System.Windows.Automation.TreeScope]::Descendants,(New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty,'参照...'))))
    if ($buttons.Count -lt 2) { throw ('Expected the folder and file browse buttons, found ' + $buttons.Count) }
    foreach ($index in 0,1) {
        $kind=if ($index -eq 0) {'folder'} else {'file'}
        InvokeNode $buttons[$index]
        $picker=$null
        WaitUntil {
            foreach ($w in TopLevel) { if ($known -notcontains $w.Current.NativeWindowHandle) { $script:picker=$w; return $true } }
            return $false
        } ($kind + ' picker did not appear')
        $pickerHwnd=[IntPtr]$picker.Current.NativeWindowHandle
        $owner=[BrowseProbe]::GetWindow($pickerHwnd,4)
        $surfaceEnabled=[BrowseProbe]::IsWindowEnabled($settingsHwnd)
        $record=[pscustomobject]@{time=(Get-Date -Format o);kind=$kind;title=$picker.Current.Name;class=[BrowseProbe]::ClassOf($pickerHwnd);owner=$owner.ToInt64();settings=$settingsHwnd.ToInt64();surfaceEnabledWhileOpen=$surfaceEnabled}
        $trace.Add($record)
        if ($owner -ne $settingsHwnd) { throw ($kind + ' picker is not owned by the settings surface (owner=' + $owner.ToInt64() + ')') }
        if ($surfaceEnabled) { throw ($kind + ' picker left the settings surface enabled underneath') }
        [BrowseProbe]::PostMessage($pickerHwnd,0x10,[IntPtr]::Zero,[IntPtr]::Zero) | Out-Null
        WaitUntil { -not (TopLevel | Where-Object { $_.Current.NativeWindowHandle -eq $pickerHwnd.ToInt64() }) } ($kind + ' picker did not close')
        WaitUntil { [BrowseProbe]::IsWindowEnabled($settingsHwnd) } ($kind + ' picker left the settings surface disabled')
        Write-Output ('PASS ' + $kind + ' picker owned by settings, surface modal, recovered')
    }
    InvokeNode (Node $settings 'name' 'キャンセル')
    WaitUntil { $main.Current.IsEnabled -and $null -eq (SettingsWindow) } 'Settings did not close after the pickers'
    Write-Output 'PASS settings closed and main window enabled'
    Write-Output 'TOTAL 3 passed; 0 failed'
} finally {
    [IO.File]::WriteAllText([IO.Path]::GetFullPath($Output),($trace | ConvertTo-Json -Depth 4),(New-Object Text.UTF8Encoding($false)))
}
