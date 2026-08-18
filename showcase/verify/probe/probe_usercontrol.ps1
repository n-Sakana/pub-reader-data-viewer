param([string]$Book, [string]$Res)
$ErrorActionPreference = 'Continue'
$out = New-Object System.Collections.ArrayList
function Note($m) { [void]$out.Add([string]$m); [System.IO.File]::WriteAllLines($Res, $out) }

$before = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$xl = $null
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.ScreenUpdating = $false
    $xl.EnableEvents = $false
    $xl.AutomationSecurity = 1
    $wb = $xl.Workbooks.Open($Book, 0, $true)
    $q = "'" + $wb.Name + "'!"
    Note ('UserControl as created -> ' + $xl.UserControl)

    $xl.UserControl = $false
    Note ('A (UserControl=False) arm  -> ' + $xl.Run($q + 'PbArmCount'))
    Start-Sleep -Seconds 5
    Note ('A (UserControl=False) count-> ' + $xl.Run($q + 'PbFiredCount'))

    $xl.UserControl = $true
    Note ('B (UserControl=True)  arm  -> ' + $xl.Run($q + 'PbArmCount'))
    Start-Sleep -Seconds 5
    Note ('B (UserControl=True)  count-> ' + $xl.Run($q + 'PbFiredCount'))
} catch {
    Note ('ERROR ' + $_.Exception.Message)
}
try { if ($xl) { $xl.DisplayAlerts = $false; $xl.Quit() } } catch {}
Start-Sleep -Milliseconds 800
Get-Process -Name EXCEL -ErrorAction SilentlyContinue |
    Where-Object { $before -notcontains $_.Id } |
    ForEach-Object { Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue }
Note 'DONE'
