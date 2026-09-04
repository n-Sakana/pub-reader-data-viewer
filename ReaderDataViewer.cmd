@echo off
setlocal EnableExtensions
set "RDV_PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%RDV_PS%" set "RDV_PS=powershell.exe"

"%RDV_PS%" -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0ReaderDataViewer.ps1" %*
set "RDV_EXIT=%ERRORLEVEL%"
if not "%RDV_EXIT%"=="0" (
  echo.
  echo Reader Data Viewer stopped with exit code %RDV_EXIT%.
  pause
)
endlocal & exit /b %RDV_EXIT%
