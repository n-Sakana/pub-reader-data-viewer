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
"%RDV_PS%" -NoProfile -ExecutionPolicy Bypass -STA -Command "$a='#RDV'+'-PS-BEGIN'; $b='#RDV'+'-CS-BEGIN'; $t=[IO.File]::ReadAllText($env:RDV_SELF,[Text.Encoding]::UTF8); $i=$t.IndexOf($a); $j=$t.IndexOf($b); if($i -lt 0 -or $j -lt 0){ Write-Host 'package markers missing'; exit 9 }; $global:RdvCs=$t.Substring($j+$b.Length); Invoke-Expression $t.Substring($i+$a.Length, $j-$i-$a.Length)"
set "RDV_RC=%ERRORLEVEL%"
endlocal & exit /b %RDV_RC%
