@echo off
setlocal EnableExtensions

rem ===== 設定（この bat と同じフォルダに ps1 / hosts.txt / password.txt を置く想定）=====
set "SCRIPT=Get-CiscoInterfaces-PerIF.ps1"
set "HOSTS=hosts.txt"
set "PASSFILE=password.txt"
set "USERNAME=cisco"

rem 繰り返し
set "REPEAT=1"
set "INTERVAL_MIN=60"
set "DURATION_MIN=0"  rem 0=無制限（REPEAT=1のとき適用）

rem 失敗時に画面を閉じない（1=有効/0=無効）
set "DEBUG_HOLD=1"
rem ======================================================================

rem --- 文字コード（日本語コンソール用 / Shift-JIS）---
chcp 932 >nul 2>&1

rem --- 実行ディレクトリへ移動（UNC/OneDrive でも pushd で安定） ---
set "BASE=%~dp0"
pushd "%BASE%" >nul 2>&1
if errorlevel 1 (
  echo [ERROR] 作業フォルダへ移動できません: "%BASE%"
  if "%DEBUG_HOLD%"=="1" pause
  exit /b 1
)

rem --- PowerShell 実行ファイルの検出（pwsh 優先） ---
set "PS=pwsh.exe"
where pwsh.exe >nul 2>&1 || set "PS=powershell.exe"

rem --- 前提ファイル確認 ---
if not exist "%SCRIPT%"  ( echo [ERROR] スクリプトが見つかりません: "%CD%\%SCRIPT%"  & goto :fail )
if not exist "%HOSTS%"   ( echo [ERROR] hosts.txt が見つかりません: "%CD%\%HOSTS%"   & goto :fail )

if exist "%PASSFILE%" (
  set "PWMSG=Password : file %CD%\%PASSFILE%"
) else (
  set "PWMSG=Password : prompt"
)

rem --- PowerShell 引数の組み立て（\" は使わない / フルパスで確実に） ---
set "PSARGS=-NoProfile -ExecutionPolicy Bypass -File ""%CD%\%SCRIPT%"" -HostsFile ""%CD%\%HOSTS%"" -Username ""%USERNAME%"""
if exist "%PASSFILE%" set "PSARGS=%PSARGS% -PasswordFile ""%CD%\%PASSFILE%"""
if "%REPEAT%"=="1"     set "PSARGS=%PSARGS% -Repeat -IntervalMinutes %INTERVAL_MIN%"
if not "%DURATION_MIN%"=="0" set "PSARGS=%PSARGS% -DurationMinutes %DURATION_MIN%"

echo * 実行開始: %date% %time%
echo   PS        : %PS%
echo   Script    : %CD%\%SCRIPT%
echo   Hosts     : %CD%\%HOSTS%
echo   Username  : %USERNAME%
echo   %PWMSG%
if "%REPEAT%"=="1" (
  echo   Repeat   : every %INTERVAL_MIN% min ^| Duration=%DURATION_MIN% min
) else (
  echo   Repeat   : no
)

echo --- 実行コマンド ---
echo %PS% %PSARGS%
echo --------------------

"%PS%" %PSARGS%
set "RC=%ERRORLEVEL%"
echo * 終了コード: %RC%
if not "%RC%"=="0" echo [ERROR] PowerShell スクリプトが異常終了しました。上のメッセージを確認してください.

goto :end

:fail
set "RC=1"

:end
popd >nul 2>&1
if "%DEBUG_HOLD%"=="1" pause
exit /b %RC%
