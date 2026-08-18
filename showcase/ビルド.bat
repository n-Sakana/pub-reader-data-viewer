@echo off
setlocal EnableExtensions
chcp 932 >nul
title VBA Pixel Bridge - ビルド
cd /d "%~dp0"

rem ---------------------------------------------------------------------------
rem  ダブルクリックで使うビルド入口。
rem  管理者権限も UAC も要らない。ExecutionPolicy はこの呼び出しの中だけ
rem  -ExecutionPolicy Bypass で外すので、利用者が設定を変える必要もない。
rem  1 / 2 / 4 のどれを選んだかで別の成果物ができる。既定値は置かない。
rem ---------------------------------------------------------------------------

rem 引数で選ぶこともできる（ビルド.bat 2）。ダブルクリックしたときは引数が
rem 無いので、下の対話メニューへ進む。引数が 1 か 2 か 4 以外なら、何も作らずに
rem 止める。ここでも黙って別の値へは倒さない。
if "%~1"=="" goto :menu
set "PBUNIT=%~1"
if "%PBUNIT%"=="1" goto :build
if "%PBUNIT%"=="2" goto :build
if "%PBUNIT%"=="4" goto :build
echo.
echo   その指定は選べません。1 か 2 か 4 にしてください。
echo.
exit /b 1

:menu
set /a PBTRY+=1
if %PBTRY% GTR 4 goto :toomany
cls
echo ==============================================================
echo   VBA Pixel Bridge - Excel分離アーキテクチャ 技術ショーケース
echo   ビルド
echo ==============================================================
echo.
echo   疑似ピクセル GUI の「1 セルが受け持つ設計ピクセル数」を選びます。
echo   小さいほど絵は細かくなり、そのぶん Excel の再描画が重くなります。
echo.
echo      1 ... 1px   仕様どおりの細かさ。盤面 812 x 812 = 659,344 セル
echo      2 ... 2px   中間。            盤面 406 x 406 = 164,836 セル
echo      4 ... 4px   いちばん軽い。    盤面 203 x 203 =  41,209 セル
echo.
echo      Q ... やめる
echo.
set "PBUNIT="
set /p "PBUNIT=  1 / 2 / 4 のどれかを入力して Enter : "

if /i "%PBUNIT%"=="Q" goto :quit
if "%PBUNIT%"=="1" goto :build
if "%PBUNIT%"=="2" goto :build
if "%PBUNIT%"=="4" goto :build

echo.
if "%PBUNIT%"=="" (
   echo   何も入力されませんでした。
) else (
   echo   "%PBUNIT%" は選べません。
)
echo   1 / 2 / 4 のどれかを入力してください。勝手に別の値では作りません。
echo.
pause
goto :menu

:build
echo.
echo   %PBUNIT%px でビルドします。しばらくお待ちください...
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0build\build_showcase.ps1" -Unit %PBUNIT%
set "RC=%ERRORLEVEL%"
echo.
if not "%RC%"=="0" (
   echo   [失敗] ビルドできませんでした。上に出ている内容を確認してください。
   echo.
   pause
   goto :menu
)
echo   [完了] dist\VBA Pixel Bridge %PBUNIT%px.xlsm
echo.
echo   このファイルをダブルクリックすると起動します。
echo   「コンテンツの有効化」を押してください。
echo   画面の下に「1 セル = %PBUNIT%px」と出ていれば、選んだとおりのものです。
echo.
pause
goto :eof

:toomany
echo.
echo   入力が選択肢と合いません。何も作らずに終わります。
echo   もう一度このファイルをダブルクリックしてやり直してください。
echo.
pause
goto :eof

:quit
echo.
echo   何も作らずに終わります。
echo.
pause
goto :eof
