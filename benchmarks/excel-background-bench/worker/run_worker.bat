@echo off
rem ===========================================================================
rem run_worker.bat -- compile (only if needed) and launch ZipWorker.exe
rem
rem This file is ASCII-only on purpose. Excel emits it verbatim at run time via
rem modZipEmit.bas, and cmd.exe reads .bat files in the OEM/ANSI code page, so
rem any non-ASCII byte here would turn into mojibake or a broken command.
rem Japanese explanation lives in README.md instead.
rem
rem   %1 = run token   %2 = window mode   %3 = parent (Excel) pid
rem
rem Compiling is skipped when ZipWorker.exe already exists, so the "prebuilt"
rem variant (Excel copies a ready-made exe into the run folder) and the
rem "emitted" variant (Excel writes ZipWorker.cs and this file, nothing else)
rem both go through exactly the same launch path.
rem ===========================================================================
setlocal
set "DIR=%~dp0"
set "TOKEN=%~1"
set "MODE=%~2"
set "PPID=%~3"
if "%TOKEN%"=="" set "TOKEN=zipbench"
if "%MODE%"=="" set "MODE=offscreen"
if "%PPID%"=="" set "PPID=0"

if exist "%DIR%ZipWorker.exe" goto :launch

set "CSC=%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if not exist "%CSC%" set "CSC=%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\csc.exe"
if not exist "%CSC%" (
  echo NO_CSC_FOUND>"%DIR%build.log"
  exit /b 9
)

rem /codepage:65001 is belt-and-braces: ZipWorker.cs is always written as UTF-8
rem with a BOM, which csc.exe honours on its own, but a stripped BOM would
rem otherwise silently change the Japanese string literals.
"%CSC%" /nologo /target:winexe /platform:anycpu /optimize+ /codepage:65001 ^
  /out:"%DIR%ZipWorker.exe" ^
  /reference:System.dll /reference:System.Drawing.dll /reference:System.Windows.Forms.dll ^
  "%DIR%ZipWorker.cs" >"%DIR%build.log" 2>&1

if not exist "%DIR%ZipWorker.exe" (
  echo BUILD_FAILED>>"%DIR%build.log"
  exit /b 1
)
echo BUILD_OK>>"%DIR%build.log"

:launch
start "" /b "%DIR%ZipWorker.exe" --token %TOKEN% --mode %MODE% --parentpid %PPID% --log "%DIR%worker.log"
exit /b 0
