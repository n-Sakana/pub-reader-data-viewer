@echo off
setlocal EnableExtensions
set "RDV_PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" set "RDV_PS=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%RDV_PS%" (
  echo Windows PowerShell 5.1 is required.
  exit /b 3
)
"%RDV_PS%" -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0tools\Build.ps1" %*
set "RDV_EXIT=%ERRORLEVEL%"
if "%~1"=="" pause
endlocal & exit /b %RDV_EXIT%
