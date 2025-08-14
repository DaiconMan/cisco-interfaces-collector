@echo off
setlocal ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION

rem ===================== 設定（必要に応じて編集） =====================
set "BASE=%~dp0"
set "SCRIPT=%BASE%Get-CiscoInterfaces-PerIF.ps1"
set "HOSTS=%BASE%hosts.txt"
set "PASSFILE=%BASE%password.txt"
set "USERNAME=cisco

rem 繰り返し設定：REPEAT=1 で有効、0 で1回のみ
set "REPEAT=1"
set "INTERVAL_MIN=60"
rem 任意：総実行時間（分）。0 なら無制限
set "DURATION_MIN=0"

rem 失敗時に画面を閉じない（1=有効/0=無効）
set "DEBUG_HOLD=1"
rem ====================================================================

rem --- PowerShell 実行ファイルの検出（pwsh 優先） ---
set "PS=pwsh.exe"
where "%PS%" >nul 2>&1
if errorlevel 1 set "PS=powershell.exe"

rem --- 前提ファイル確認 ---
if not exist "%SCRIPT%" (
  echo [ERROR] スクリプトが見つかりません: "%SCRIPT%"
  if "%DEBUG_HOLD%"=="1" pause
  exit /b 1
)
if not exist "%HOSTS%" (
  echo [ERROR] hosts.txt が見つかりません: "%HOSTS%"
  if "%DEBUG_HOLD%"=="1" pause
  exit /b 1
)

rem --- オプションの組み立て（引用符は素直に） ---
set "PWOPT="
if exist "%PASSFILE%" set "PWOPT=-PasswordFile \"%PASSFILE%\""

set "REPEATOPT="
if "%REPEAT%"=="1" (
  set "REPEATOPT=-Repeat -IntervalMinutes %INTERVAL_MIN%"
  if not "%DURATION_MIN%"=="0" set "REPEATOPT=%REPEATOPT% -DurationMinutes %DURATION_MIN%"
)

echo * 実行開始: %DATE% %TIME%
echo   PS        : %PS%
echo   Script    : %SCRIPT%
echo   Hosts     : %HOSTS%
echo   Username  : %USERNAME%
if defined PWOPT (echo   Password : file %PASSFILE%) else (echo   Password : prompt)
if defined REPEATOPT (echo   Repeat   : every %INTERVAL_MIN% min ^| Duration=%DURATION_MIN% min) else (echo   Repeat   : no)

"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -HostsFile "%HOSTS%" -Username "%USERNAME%" %PWOPT% %REPEATOPT%
set "RC=%ERRORLEVEL%"

echo * 終了コード: %RC%
if not "%RC%"=="0" (
  echo [ERROR] PowerShell スクリプトが異常終了しました。上のメッセージを確認してください。
  if "%DEBUG_HOLD%"=="1" pause
)
endlocal & exit /b %RC%
