@echo off
setlocal EnableExtensions

rem ===== 設定（この bat と同じフォルダに PS1 と logs/ を置く想定）=====
set "SCRIPT=Build-InterfacesReport.ps1"
set "LOGS=logs"
set "OUTFILE=interfaces_report.html"
set "TOPN=10"
set "UNUSED_THRESHOLD=0"

rem ── 閾値（Teams/SharePoint 調査向けの推奨値）
set "UTIL_WARN_PCT=40"
set "UTIL_SEVERE_PCT=70"
set "PPS_WARN=50000"

rem ── 表示名マッピング（IP,表示名 のCSV。無い場合は自動スキップ）
set "HOSTS=hosts.txt"

set "VERBOSE=0"     rem 1=詳細表示ON
set "DEBUG_HOLD=1"  rem 1=終了時に停止
rem ===============================================================

chcp 932 >nul 2>&1

set "BASE=%~dp0"
pushd "%BASE%" >nul 2>&1
if errorlevel 1 (
  echo [ERROR] 作業フォルダへ移動できません: "%BASE%"
  if "%DEBUG_HOLD%"=="1" pause
  exit /b 1
)

for %%Z in ("%SCRIPT%") do set "EXT=%%~xZ"
if /I not "%EXT%"==".ps1" set "SCRIPT=%SCRIPT%.ps1"

for %%I in ("%SCRIPT%")  do set "ABS_SCRIPT=%%~fI"
for %%I in ("%LOGS%")    do set "ABS_LOGS=%%~fI"
for %%I in ("%OUTFILE%") do set "ABS_OUT=%%~fI"
for %%I in ("%HOSTS%")   do set "ABS_HOSTS=%%~fI"

if not exist "%ABS_SCRIPT%" ( echo [ERROR] スクリプトが見つかりません: "%ABS_SCRIPT%" & goto :fail )
if not exist "%ABS_LOGS%"   ( echo [ERROR] ログフォルダが見つかりません: "%ABS_LOGS%"  & goto :fail )

set "PS=pwsh.exe"
where pwsh.exe >nul 2>&1 || set "PS=powershell.exe"

set "EXTRA="
if "%VERBOSE%"=="1" set "EXTRA=-Verbose"

set "HOSTS_ARG="
if exist "%ABS_HOSTS%" set "HOSTS_ARG=-HostsFile ""%ABS_HOSTS%"""

echo * 実行開始: %date% %time%
echo   PS       : %PS%
echo   Script   : %ABS_SCRIPT%
echo   LogsRoot : %ABS_LOGS%
echo   OutFile  : %ABS_OUT%
echo   TopN     : %TOPN%
echo   UnusedTh : %UNUSED_THRESHOLD%
echo   UtilWarn : %UTIL_WARN_PCT%%%
echo   UtilSevr : %UTIL_SEVERE_PCT%%%
echo   PPSWarn  : %PPS_WARN% pps
if exist "%ABS_HOSTS%" echo   HostsMap : %ABS_HOSTS%
if "%VERBOSE%"=="1" echo   Verbose  : on

echo --- 実行コマンド ---
echo "%PS%" -NoProfile -ExecutionPolicy Bypass -File "%ABS_SCRIPT%" -LogsRoot "%ABS_LOGS%" -OutFile "%ABS_OUT%" -TopN %TOPN% -UnusedThreshold %UNUSED_THRESHOLD% -UtilWarnPct %UTIL_WARN_PCT% -UtilSeverePct %UTIL_SEVERE_PCT% -PpsWarn %PPS_WARN% %HOSTS_ARG% %EXTRA%
echo --------------------

"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%ABS_SCRIPT%" -LogsRoot "%ABS_LOGS%" -OutFile "%ABS_OUT%" -TopN %TOPN% -UnusedThreshold %UNUSED_THRESHOLD% -UtilWarnPct %UTIL_WARN_PCT% -UtilSeverePct %UTIL_SEVERE_PCT% -PpsWarn %PPS_WARN% %HOSTS_ARG% %EXTRA%
set "RC=%ERRORLEVEL%"
echo * 終了コード: %RC%
if not "%RC%"=="0" echo [ERROR] 生成に失敗しました。上のメッセージを確認してください.
goto :end

:fail
set "RC=1"

:end
popd >nul 2>&1
if "%DEBUG_HOLD%"=="1" pause
exit /b %RC%