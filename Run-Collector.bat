@echo off
setlocal EnableExtensions EnableDelayedExpansion

rem ===================== 設定 =====================
set "BASE=%~dp0"
set "SCRIPT=%BASE%Get-CiscoInterfaces-PerIF.ps1"
set "HOSTS=%BASE%hosts.txt"
set "PASSFILE=%BASE%password.txt"
set "USERNAME=cisco"

rem 繰り返し: 1=有効, 0=1回のみ
set "REPEAT=1"
set "INTERVAL_MIN=60"
rem 総実行時間(分)。0=無制限（REPEAT=1の時のみ有効）
set "DURATION_MIN=0"

rem 失敗時に画面を閉じない: 1=有効/0=無効
set "DEBUG_HOLD=1"
rem =================================================

rem --- PowerShell 実行ファイルの検出（pwsh 優先） ---
set "PS=pwsh.exe"
where pwsh.exe >nul 2>&1 || set "PS=powershell.exe"

rem --- 前提ファイル確認 ---
if not exist "%SCRIPT%" (
  echo [ERROR] スクリプトが見つかりません: "%SCRIPT%"
  goto :fail
)
if not exist "%HOSTS%" (
  echo [ERROR] hosts.txt が見つかりません: "%HOSTS%"
  goto :fail
)

rem 表示用メッセージ
if exist "%PASSFILE%" (
  set "PWMSG=Password : file %PASSFILE%"
) else (
  set "PWMSG=Password : prompt"
)

rem --- PowerShell 引数を正しく組み立て（\" は使わない） ---
set "PSARGS=-NoProfile -ExecutionPolicy Bypass -File ""%SCRIPT%"" -HostsFile ""%HOSTS%"" -Username ""%USERNAME%"""
if exist "%PASSFILE%" set "PSARGS=%PSARGS% -PasswordFile ""%PASSFILE%"""
if "%REPEAT%"=="1" set "PSARGS=%PSARGS% -Repeat -IntervalMinutes %INTERVAL_MIN%"
if not "%DURATION_MIN%"=="0" set "PSARGS=%PSARGS% -DurationMinutes %DURATION_MIN%"

echo * 実行開始: %date% %time%
echo   PS        : %PS%
echo   Script    : %SCRIPT%
echo   Hosts     : %HOSTS%
echo   Username  : %USERNAME%
echo   %PWMSG%
if "%REPEAT%"=="1" (
  echo   Repeat   : every %INTERVAL_MIN% min ^| Duration=%DURATION_MIN% min
) else (
  echo   Repeat   : no
)

"%PS%" %PSARGS%
set "RC=%ERRORLEVEL%"
echo * 終了コード: %RC%
if not "%RC%"=="0" echo [ERROR] PowerShell スクリプトが異常終了しました。上のメッセージを確認してください.

goto :end

:fail
set "RC=1"

:end
if "%DEBUG_HOLD%"=="1" pause
exit /b %RC%
