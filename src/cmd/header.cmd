@echo off
setlocal EnableExtensions
rem ==========================================================================
rem  @@TITLE@@
rem
rem  Reader Data Viewer -- single file. There is nothing to install and nothing
rem  to unpack: this .cmd starts Windows PowerShell at normal privilege, pulls
rem  the PowerShell bootstrap and the C# program out of its own tail, compiles
rem  the C# with the .NET Framework compiler that ships with Windows, and runs
rem  it. No .ps1, .cs, .exe or .dll is needed beside it.
rem
rem  usage:  @@USAGE@@
rem          the data directory defaults to ..\data then .\data
rem ==========================================================================
set "RDV_SELF=%~f0"
set "RDV_ARGS=%*"
set "RDV_PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%RDV_PS%" set "RDV_PS=powershell.exe"

rem  --- no console window ---------------------------------------------------
rem  A .cmd started from Explorer always gets a console, and on Windows 11 that
rem  console is hosted by Windows Terminal, where hiding it afterwards does
rem  nothing: GetConsoleWindow returns the pseudo console, not the terminal
rem  window. So the first run starts ITSELF again with CreateNoWindow -- the
rem  child gets no console at all -- and then exits, which closes this one.
rem  What the operator sees is the app's window and nothing else.
if not defined RDV_NOWIN (
  set "RDV_NOWIN=1"
  "%RDV_PS%" -NoProfile -ExecutionPolicy Bypass -Command "$si = New-Object Diagnostics.ProcessStartInfo; $si.FileName = $env:ComSpec; $si.Arguments = '/c \"' + $env:RDV_SELF + '\" ' + $env:RDV_ARGS; $si.UseShellExecute = $false; $si.CreateNoWindow = $true; $si.WorkingDirectory = (Split-Path -Parent $env:RDV_SELF); [void][Diagnostics.Process]::Start($si)"
  endlocal & exit /b 0
)
"%RDV_PS%" -NoProfile -ExecutionPolicy Bypass -STA -Command "$a='#RDV'+'-PS-BEGIN'; $b='#RDV'+'-CS-BEGIN'; $t=[IO.File]::ReadAllText($env:RDV_SELF,[Text.Encoding]::UTF8); $i=$t.IndexOf($a); $j=$t.IndexOf($b); if($i -lt 0 -or $j -lt 0){ Write-Host 'package markers missing'; exit 9 }; $global:RdvCs=$t.Substring($j+$b.Length); Invoke-Expression $t.Substring($i+$a.Length, $j-$i-$a.Length)"
set "RDV_RC=%ERRORLEVEL%"
endlocal & exit /b %RDV_RC%
