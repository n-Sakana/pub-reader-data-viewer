@echo off
rem ==========================================================================
rem  build.bat -- double-click this. It builds the C# product in dist\
rem  from the sources in this repository.
rem
rem  It does NOT need administrator rights and never asks for elevation.
rem  It does NOT write to the registry.
rem  It does NOT change the machine's execution policy: -ExecutionPolicy
rem  Bypass below applies to the one PowerShell process it starts.
rem  It does NOT need or start Excel.
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

rem  Every %VAR% below is echoed INSIDE QUOTES on purpose. cmd expands a
rem  variable before it parses the line, so a folder called
rem  "C:\Tools (x86) & co" turns an unquoted echo into a broken command --
rem  and inside an if(...) block a ')' from the path closes the block, which
rem  aborts the whole file before it reaches the pause at the end. That is
rem  the "double-click, window vanishes, nothing built" failure.
echo ==========================================================================
echo  Reader Data Viewer -- build
echo  repository : "%RDV_ROOT%"
echo  powershell : "%RDV_PS%"
echo  builds     : dist\app-csharp
echo ==========================================================================
echo.

if not exist "%RDV_SCRIPT%" (
  echo [ERROR] the build script was not found:
  echo         "%RDV_SCRIPT%"
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
  echo  The product is under:
  echo      "%RDV_ROOT%\dist\app-csharp"
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
