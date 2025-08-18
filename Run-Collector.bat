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

rem --- 日本語コンソール向け（Shift-JIS） ---
chcp 932 >nul 2>&1

rem --- 実行ディレクトリへ移動（UNC/OneDrive でも pushd で安定） ---
set "BASE=%~dp0"
pushd "%BASE%" >nul 2>&1
if errorlevel 1 (
  echo [ERROR] 作業フォルダへ移動できません: "%BASE%"
  if "%DEBUG_HOLD%"=="1" pause
  exit /b 1
)

rem --- 拡張子が無ければ .ps1 を付与（誤設定対策） ---
for %%Z in ("%SCRIPT%") do (
  set "EXT=%%~xZ"
)
if /I not "%EXT%"==".ps1" set "SCRIPT=%SCRIPT%.ps1"

rem --- 絶対パスを解決（スペース/日本語/UNC 安定化） ---
for %%I in ("%SCRIPT%")  do set "ABS_SCRIPT=%%~fI"
for %%I in ("%HOSTS%")   do set "ABS_HOSTS=%%~fI"
if exist "%PASSFILE%" (
  for %%I in ("%PASSFILE%") do set "ABS_PASSFILE=%%~fI"
) else (
  set "ABS_PASSFILE="
)

rem --- 前提ファイル確認 ---
if not exist "%ABS_SCRIPT%" ( echo [ERROR] スクリプトが見つかりません: "%ABS_SCRIPT%" & goto :fail )
if not exist "%ABS_HOSTS%"  ( echo [ERROR] hosts.txt が見つかりません: "%ABS_HOSTS%"  & goto :fail )

if defined ABS_PASSFILE (
  set "PWMSG=Password : file %ABS_PASSFILE%"
) else (
  set "PWMSG=Password : prompt"
)

rem --- 追加オプション（繰り返し） ---
set "REPEAT_OPTS="
if "%REPEAT%"=="1" set "REPEAT_OPTS=%REPEAT_OPTS% -Repeat -IntervalMinutes %INTERVAL_MIN%"
if not "%DURATION_MIN%"=="0" set "REPEAT_OPTS=%REPEAT_OPTS% -DurationMinutes %DURATION_MIN%"

rem --- PowerShell 実行ファイルの検出（pwsh 優先） ---
set "PS=pwsh.exe"
where pwsh.exe >nul 2>&1 || set "PS=powershell.exe"

echo * 実行開始: %date% %time%
echo   PS        : %PS%
echo   Script    : %ABS_SCRIPT%
echo   Hosts     : %ABS_HOSTS%
echo   Username  : %USERNAME%
echo   %PWMSG%
if "%REPEAT%"=="1" (
  echo   Repeat   : every %INTERVAL_MIN% min ^| Duration=%DURATION_MIN% min
) else (
  echo   Repeat   : no
)

echo --- 実行コマンド ---
if defined ABS_PASSFILE (
  echo "%PS%" -NoProfile -ExecutionPolicy Bypass -File "%ABS_SCRIPT%" -HostsFile "%ABS_HOSTS%" -Username "%USERNAME%" -PasswordFile "%ABS_PASSFILE%" %REPEAT_OPTS%
) else (
  echo "%PS%" -NoProfile -ExecutionPolicy Bypass -File "%ABS_SCRIPT%" -HostsFile "%ABS_HOSTS%" -Username "%USERNAME%" %REPEAT_OPTS%
)
echo --------------------

rem --- 実行（-File には必ず .ps1 の絶対パスを渡す） ---
if defined ABS_PASSFILE (
  "%PS%" -NoProfile -ExecutionPolicy Bypass -File "%ABS_SCRIPT%" -HostsFile "%ABS_HOSTS%" -Username "%USERNAME%" -PasswordFile "%ABS_PASSFILE%" %REPEAT_OPTS%
) else (
  "%PS%" -NoProfile -ExecutionPolicy Bypass -File "%ABS_SCRIPT%" -HostsFile "%ABS_HOSTS%" -Username "%USERNAME%" %REPEAT_OPTS%
)

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
