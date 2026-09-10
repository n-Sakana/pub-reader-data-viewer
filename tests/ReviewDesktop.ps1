param([Parameter(Mandatory=$true)][int]$Owner, [long]$Handle=0,
    [ValidateSet('inspect','click','type','keys','chord','capture','move')][string]$Action='inspect',
    [string]$Id='', [string]$Name='', [string]$Type='', [string]$Text='', [string]$Output='', [int]$Index=0,
    [int]$X=50,[int]$Y=50,[int]$Width=1300,[int]$Height=980)
$ErrorActionPreference='Stop'
Add-Type -AssemblyName UIAutomationClient,UIAutomationTypes,WindowsBase,System.Drawing,System.Windows.Forms
Add-Type -Path (Join-Path $PSScriptRoot 'ReviewDesktop.cs') -ReferencedAssemblies UIAutomationClient,UIAutomationTypes,WindowsBase,System.Drawing,System.Windows.Forms
[ReviewDesktop]::Physical()
$windows=@([ReviewDesktop]::Windows($Owner))
if(-not $Handle){
    foreach($hwnd in $windows){$e=[Windows.Automation.AutomationElement]::FromHandle([IntPtr]$hwnd);[pscustomobject]@{handle=$hwnd;pid=$Owner;name=$e.Current.Name;bounds=$e.Current.BoundingRectangle.ToString();enabled=$e.Current.IsEnabled}}
    return
}
if(-not [ReviewDesktop]::IsOwnedVisible([IntPtr]$Handle,$Owner)){throw 'Not a visible window of the specified test PID'}
$root=[Windows.Automation.AutomationElement]::FromHandle([IntPtr]$Handle)
if($Action -eq 'capture'){[ReviewDesktop]::Capture([IntPtr]$Handle,[IO.Path]::GetFullPath($Output));return}
if($Action -eq 'move'){[ReviewDesktop]::Move([IntPtr]$Handle,$X,$Y,$Width,$Height);return}
if($Action -eq 'keys'){[ReviewDesktop]::Focus([IntPtr]$Handle);[Windows.Forms.SendKeys]::SendWait($Text);return}
if($Action -eq 'chord'){[ReviewDesktop]::Focus([IntPtr]$Handle);[ReviewDesktop]::Chord([byte[]]($Text.Split(',')|ForEach-Object{[byte]$_}));return}
$nodes=$root.FindAll([Windows.Automation.TreeScope]::Descendants,[Windows.Automation.Condition]::TrueCondition)
if($Action -eq 'inspect'){
    foreach($node in $nodes){$p=$node.Current;if(-not $p.IsOffscreen -and $p.BoundingRectangle.Width -gt 0){$value='';$pattern=$null;if($node.TryGetCurrentPattern([Windows.Automation.ValuePattern]::Pattern,[ref]$pattern)){$value=$pattern.Current.Value};[pscustomobject]@{id=$p.AutomationId;name=$p.Name;type=$p.ControlType.ProgrammaticName;enabled=$p.IsEnabled;value=$value;bounds=$p.BoundingRectangle.ToString()}}}
    return
}
$found=@($nodes|Where-Object{(-not $Id -or $_.Current.AutomationId -eq $Id) -and (-not $Name -or $_.Current.Name -eq $Name) -and (-not $Type -or $_.Current.ControlType.ProgrammaticName -eq ('ControlType.'+$Type)) -and -not $_.Current.IsOffscreen})
if($found.Count -le $Index){throw 'Visible control not found'}
if($found[$Index].Current.IsKeyboardFocusable){$found[$Index].SetFocus()}
[ReviewDesktop]::Focus([IntPtr]$Handle)
[ReviewDesktop]::Click($found[$Index])
if($Action -eq 'type'){
    [Windows.Forms.SendKeys]::SendWait('^a')
    [Windows.Forms.Clipboard]::SetText($Text)
    [Windows.Forms.SendKeys]::SendWait('^v')
}
Start-Sleep -Milliseconds 150
