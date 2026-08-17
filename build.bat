@echo off
rem ==========================================================================
rem  build.bat -- double-click this. It builds EVERY distributable in dist\
rem  from the sources in this repository (C# and VBA alike).
rem
rem  It does NOT need administrator rights and never asks for elevation.
rem  It does NOT write to the registry.
rem  It does NOT change the machine's execution policy: -ExecutionPolicy
rem  Bypass below applies to the one PowerShell process it starts.
rem  It does NOT touch any Excel instance it did not start itself.
rem
rem  It DOES need Excel, plus Excel's per-user setting
rem      Trust Center > Macro Settings > Trust access to the VBA project
rem      object model
rem  because the only way this repository can put VBA into an .xlsm is
rem  VBProject.VBComponents.Import. That setting is per user (HKCU) and needs
rem  no administrator; if it is off, the build stops and prints how to turn it
rem  on, instead of quietly producing fewer files.
rem
rem  The window stays open at the end -- on success and on failure alike -- so
rem  the reason and the exit code can be read.
rem ==========================================================================
setlocal EnableExtensions

set "RDV_ROOT=%~dp0"
if "%RDV_ROOT:~-1%"=="\" set "RDV_ROOT=%RDV_ROOT:~0,-1%"

set "RDV_PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%RDV_PS%" set "RDV_PS=powershell.exe"

set "RDV_SCRIPT=%RDV_ROOT%\build\build_dist.ps1"
set "RDV_LOGDIR=%RDV_ROOT%\work"

echo ==========================================================================
echo  Reader Data Viewer -- build
echo  repository : %RDV_ROOT%
echo  powershell : %RDV_PS%
echo  builds     : the product only -- dist\app-csharp + dist\app-vba
echo ==========================================================================
echo.

if not exist "%RDV_SCRIPT%" (
  echo [ERROR] the build script was not found:
  echo         %RDV_SCRIPT%
  echo.
  echo         Run build.bat from the folder it lives in, inside a complete
  echo         checkout of the repository.
  set "RDV_RC=9"
  goto :report
)

if not exist "%RDV_LOGDIR%" mkdir "%RDV_LOGDIR%" 2>nul

rem  -NoProfile   : the user's PowerShell profile cannot change the outcome
rem  -ExecutionPolicy Bypass : this process only; nothing is written anywhere
"%RDV_PS%" -NoProfile -ExecutionPolicy Bypass -File "%RDV_SCRIPT%" -Root "%RDV_ROOT%" 2>&1
set "RDV_RC=%ERRORLEVEL%"

:report
echo.
echo ==========================================================================
if "%RDV_RC%"=="0" (
  echo  RESULT: success   ^(exit code %RDV_RC%^)
  echo.
  echo  The products are under:
  echo      %RDV_ROOT%\dist\app-csharp
  echo      %RDV_ROOT%\dist\app-vba
) else (
  echo  RESULT: FAILED   ^(exit code %RDV_RC%^)
  echo.
  echo  The reason is printed above, between "=== FAILED" and this line,
  echo  or listed under "=== result" as the artifacts that are missing.
  echo  Nothing was installed and nothing was changed outside this folder.
)
echo ==========================================================================
echo.
echo  This window will stay open. Press any key to close it.
pause >nul
endlocal
exit /b %RDV_RC%
