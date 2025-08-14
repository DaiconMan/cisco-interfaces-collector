---

# Run-Collector-Admin.bat

> 管理者権限で自動昇格し、**ExecutionPolicy をプロセス限定で Bypass**して PowerShell スクリプトを起動します。  
> 既定は 1時間おきの無限ループで実行。必要に応じて先頭の設定値を編集してください。

```bat
@echo off
setlocal ENABLEDELAYEDEXPANSION

rem ===================== 設定（必要に応じて編集） =====================
set "SCRIPT=%~dp0Get-CiscoInterfaces-PerIF.ps1"
set "HOSTS=%~dp0hosts.txt"
set "PASSFILE=%~dp0password.txt"
set "USERNAME=commonuser"

rem 繰り返し設定：REPEAT=1 で有効、0 で1回のみ
set "REPEAT=1"
set "INTERVAL_MIN=60"
rem 任意：総実行時間（分）。0 なら無制限
set "DURATION_MIN=0"
rem ====================================================================

rem --- 管理者権限チェック（昇格して再実行） ---
>nul 2>&1 net session
if not "%errorlevel%"=="0" (
  echo * 管理者権限で再実行します...
  powershell -NoProfile -Command "Start-Process -FilePath '\"%~f0\"' -Verb RunAs"
  exit /b
)

rem --- 前提ファイル確認 ---
if not exist "%SCRIPT%" (
  echo [ERROR] スクリプトが見つかりません: "%SCRIPT%"
  exit /b 1
)
if not exist "%HOSTS%" (
  echo [ERROR] hosts.txt が見つかりません: "%HOSTS%"
  exit /b 1
)
if not exist "%PASSFILE%" (
  echo [WARN ] password.txt が見つかりません。起動後にパスワード入力を求められます。
)

rem --- PowerShell 引数を組み立て ---
set "PSARGS=-NoProfile -ExecutionPolicy Bypass -File \"%SCRIPT%\" -HostsFile \"%HOSTS%\" -Username \"%USERNAME%\""
if exist "%PASSFILE%" (
  set "PSARGS=%PSARGS% -PasswordFile \"%PASSFILE%\""
)
if "%REPEAT%"=="1" (
  set "PSARGS=%PSARGS% -Repeat -IntervalMinutes %INTERVAL_MIN%"
  if not "%DURATION_MIN%"=="0" (
    set "PSARGS=%PSARGS% -DurationMinutes %DURATION_MIN%"
  )
)

echo * 実行開始: %DATE% %TIME%
echo   Script   : %SCRIPT%
echo   Hosts    : %HOSTS%
echo   Username : %USERNAME%
if exist "%PASSFILE%" (echo   Password : (file) %PASSFILE%) else (echo   Password : (prompt))
if "%REPEAT%"=="1" (
  echo   Repeat  : every %INTERVAL_MIN% min  (Duration=%DURATION_MIN% min ^(0=unlimited^))
) else (
  echo   Repeat  : no
)

rem --- 実行 ---
powershell %PSARGS%
set "RC=%ERRORLEVEL%"

echo * 終了コード: %RC%
endlocal & exit /b %RC%