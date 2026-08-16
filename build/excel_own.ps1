# ============================================================================
# excel_own.ps1 -- WHICH Excel this script started, and may therefore end.
#
#   . (Join-Path $PSScriptRoot 'excel_own.ps1')
#   $own = New-OwnedExcel            # creates it, or THROWS
#   $xl  = $own.App
#   ...
#   try { $xl.Quit() } catch { }
#   Stop-ExcelOwned $own.Pid
#
# THE RULE: this script touches an Excel process only if this script started it.
#
# Ownership used to be decided by DIFFERENCE: snapshot Get-Process EXCEL, create
# the COM instance, snapshot again, treat everything new as mine. Any Excel that
# started inside that window -- one the operator opened -- fell into the
# difference and was force-killed on the way out.
#
# Asking the instance which process it is (Application.Hwnd, which exists even
# when Visible is false, resolved through UI Automation) is exact. No window
# enumeration, no Win32, no guessing -- the same rule the product follows.
#
# But exact is not sufficient. Excel does not always start a NEW process for a
# new Application object; it can hand back one that is already running, which is
# the operator's Excel with the operator's workbooks in it. Owning nothing in
# that case is still not enough, because the caller would go on to set Visible,
# open books in it, and Quit it at the end. So the identity is settled BEFORE
# anything is done to the instance, and a reused or unidentifiable one is a
# hard, explicit failure: the reference is released and nothing else happens.
#
# Quit is not always obeyed (a COM Excel that has ever received a real mouse
# click ignores it -- measured), which is why ending the process is still
# needed. It is now the OWN process and only that one.
# ============================================================================

function Get-PidFromHwnd([long] $hwnd) {
  # Write-Warning, never Write-Output: anything written to the output stream
  # inside a function becomes part of its RETURN VALUE, and this one returns a
  # process id that other code compares and kills on.
  if ($hwnd -eq 0) { return 0 }
  try {
    Add-Type -AssemblyName UIAutomationClient -ErrorAction Stop
    Add-Type -AssemblyName UIAutomationTypes -ErrorAction Stop
    $e = [System.Windows.Automation.AutomationElement]::FromHandle([IntPtr]$hwnd)
    if ($null -eq $e) { throw 'no automation element for that window' }
    return [int]$e.Current.ProcessId
  } catch {
    Write-Warning ("cannot identify the process behind hwnd {0}: {1}" -f $hwnd, $_.Exception.Message)
    return 0
  }
}

# Start an Excel this script owns. Returns @{ App; Pid }, or throws without
# having touched anything. NOTHING may be done to the instance before this
# returns -- that is the whole point of doing the creation here.
function New-OwnedExcel {
  $known = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
  $xl = New-Object -ComObject Excel.Application
  # a just-started Excel can be busy (0x800AC472); the handle is the first thing
  # asked of it and the only thing asked before ownership is settled
  $h = 0
  for ($i = 0; $i -lt 25; $i++) {
    try { $h = [long]$xl.Hwnd; break } catch { Start-Sleep -Milliseconds 200 }
  }
  $p = Get-PidFromHwnd $h
  if ($p -le 0 -or (@($known) -contains $p)) {
    # release the reference and do NOTHING else: no property write, no Quit, no
    # Stop-Process. If Excel handed back a running instance, that instance
    # belongs to whoever opened it.
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) } catch { }
    if ($p -gt 0) {
      # the parentheses matter: -f binds to the LAST string of a concatenation,
      # so without them the {0} in the FIRST one is never filled and the refusal
      # loses the one fact worth logging
      throw (("Excel handed back a process that was already running (pid {0}). " +
              "This script only ever drives an Excel it started itself. " +
              "Close the running Excel and run this again.") -f $p)
    }
    throw 'cannot identify the Excel process this script just started; refusing to drive it'
  }
  return [pscustomobject]@{ App = $xl; Pid = $p }
}

function Stop-ExcelOwned([int] $ownPid) {
  if ($ownPid -le 0) { return }
  $p = Get-Process -Id $ownPid -ErrorAction SilentlyContinue
  if ($null -eq $p) { return }
  if ($p.ProcessName -ne 'EXCEL') { return }     # a recycled pid is not ours
  Stop-Process -Id $ownPid -Force -ErrorAction SilentlyContinue
}
