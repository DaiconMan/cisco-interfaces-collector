@echo off
setlocal EnableExtensions EnableDelayedExpansion

rem ===================== 設定（必要に応じて変更） =====================
rem ※ この .bat と .ps1 / hosts.txt / password.txt は同じフォルダに置く前提
set "SCRIPT=Get-CiscoInterfaces-PerIF.ps1"
set "HOSTS=hosts.txt"
set "PASSFILE=password.txt"
set "USERNAME=cisco"

rem 繰り返し設定
set "REPEAT=1"
set "INTERVAL_MIN=60"
set "DURATION_MIN=0"   rem 0=無制限（REPEAT=1のときだけ有効）

rem 失敗時に画面を閉じない（1=有効/0=無効）
set "DEBUG_HOLD=1"
rem =====================================================================

rem --- 実行ディレクトリへ移動（UNCなら一時ドライブ割当） ---
set "BASE=%~dp0"
pushd "%BASE%" >nul 2>&1
if errorlevel 1 (
  echo [ERROR] 作業フォルダへ移動できません: "%BASE%"
  if "%DEBUG_HOLD%"=="1" pause
  exit /b 1
)

rem --- コンソールをUTF-8に（表示安定用。失敗しても無視） ---
chcp 65001 >nul 2>&1

rem --- PowerShell 実行ファイルの検出（pwsh 優先） ---
set "PS=pwsh.exe"
where pwsh.exe >nul 2>&1 || set "PS=powershell.exe"

rem --- 前提ファイル確認（相対パスでOK） ---
if not exist ".\%SCRIPT%"  ( echo [ERROR] スクリプトが見つかりません: ".\%SCRIPT%"  & goto :fail )
if not exist ".\%HOSTS%"   ( echo [ERROR] hosts.txt が見つかりません: ".\%HOSTS%"   & goto :fail )
set "PWMSG=Password : prompt"
if exist ".\%PASSFILE%" set "PWMSG=Password : file .\%PASSFILE%"

rem --- PowerShell 引数を相対パスで組み立て（\" は使わない） ---
set "PSARGS=-NoProfile -ExecutionPolicy Bypass -File "".\%SCRIPT%"" -HostsFile "".\%HOSTS%"" -Username ""%USERNAME%"""
if exist ".\%PASSFILE%" set "PSARGS=%PSARGS% -PasswordFile "".\%PASSFILE%"""
if "%REPEAT%"=="1"       set "PSARGS=%PSARGS% -Repeat -IntervalMinutes %INTERVAL_MIN%"
if not "%DURATION_MIN%"=="0" set "PSARGS=%PSARGS% -DurationMinutes %DURATION_MIN%"

echo * 実行開始: %date% %time%
echo   PS        : %PS%
echo   Script    : .\%SCRIPT%
echo   Hosts     : .\%HOSTS%
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
popd >nul 2>&1
if "%DEBUG_HOLD%"=="1" pause
exit /b %RC%
