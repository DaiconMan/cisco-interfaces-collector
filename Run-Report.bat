
@echo off
setlocal EnableExtensions

rem ===== 設定（この bat と同じフォルダに PS1 と logs/ を置く想定）=====
set "SCRIPT=Build-InterfacesReport.ps1"
set "LOGS=logs"
set "OUTFILE=interfaces_report.html"
set "TOPN=10"
set "UNUSED_THRESHOLD=0"
set "VERBOSE=0"     rem 1=詳細表示ON
set "DEBUG_HOLD=1"  rem 1=終了時に停止
rem ===============================================================

rem 日本語コンソールの文字化け対策（Shift-JIS）
chcp 932 >nul 2>&1

rem 実行フォルダへ移動（UNC/OneDrive でも pushd で安定）
set "BASE=%~dp0"
pushd "%BASE%" >nul 2>&1
if errorlevel 1 (
  echo [ERROR] 作業フォルダへ移動できません: "%BASE%"
  if "%DEBUG_HOLD%"=="1" pause
  exit /b 1
)

rem 拡張子が無ければ .ps1 を付与（誤設定対策）
for %%Z in ("%SCRIPT%") do set "EXT=%%~xZ"
if /I not "%EXT%"==".ps1" set "SCRIPT=%SCRIPT%.ps1"

rem 絶対パスへ解決（スペース/日本語/UNC 安定化）
for %%I in ("%SCRIPT%")  do set "ABS_SCRIPT=%%~fI"
for %%I in ("%LOGS%")    do set "ABS_LOGS=%%~fI"
for %%I in ("%OUTFILE%") do set "ABS_OUT=%%~fI"

rem 前提確認
if not exist "%ABS_SCRIPT%" ( echo [ERROR] スクリプトが見つかりません: "%ABS_SCRIPT%" & goto :fail )
if not exist "%ABS_LOGS%"   ( echo [ERROR] ログフォルダが見つかりません: "%ABS_LOGS%"  & goto :fail )

rem PowerShell 実行ファイルの検出（pwsh 優先）
set "PS=pwsh.exe"
where pwsh.exe >nul 2>&1 || set "PS=powershell.exe"

rem --- PowerShell 引数の組み立て（Run-Collector と同型／\" は使わない） ---
set "PSARGS=-NoProfile -ExecutionPolicy Bypass -File ""%ABS_SCRIPT%"" -LogsRoot ""%ABS_LOGS%"" -OutFile ""%ABS_OUT%"" -TopN %TOPN% -UnusedThreshold %UNUSED_THRESHOLD%"
if "%VERBOSE%"=="1" set "PSARGS=%PSARGS% -Verbose"

echo * 実行開始: %date% %time%
echo   PS       : %PS%
echo   Script   : %ABS_SCRIPT%
echo   LogsRoot : %ABS_LOGS%
echo   OutFile  : %ABS_OUT%
echo   TopN     : %TOPN%
echo   UnusedTh : %UNUSED_THRESHOLD%
if "%VERBOSE%"=="1" echo   Verbose  : on

"%PS%" %PSARGS%
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