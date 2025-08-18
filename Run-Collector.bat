@echo off
setlocal EnableExtensions

rem ===== �ݒ�i���� bat �Ɠ����t�H���_�� ps1 / hosts.txt / password.txt ��u���z��j=====
set "SCRIPT=Get-CiscoInterfaces-PerIF.ps1"
set "HOSTS=hosts.txt"
set "PASSFILE=password.txt"
set "USERNAME=cisco"

rem �J��Ԃ�
set "REPEAT=1"
set "INTERVAL_MIN=60"
set "DURATION_MIN=0"  rem 0=�������iREPEAT=1�̂Ƃ��K�p�j

rem ���s���ɉ�ʂ���Ȃ��i1=�L��/0=�����j
set "DEBUG_HOLD=1"
rem ======================================================================

rem --- ���{��R���\�[�������iShift-JIS�j ---
chcp 932 >nul 2>&1

rem --- ���s�f�B���N�g���ֈړ��iUNC/OneDrive �ł� pushd �ň���j ---
set "BASE=%~dp0"
pushd "%BASE%" >nul 2>&1
if errorlevel 1 (
  echo [ERROR] ��ƃt�H���_�ֈړ��ł��܂���: "%BASE%"
  if "%DEBUG_HOLD%"=="1" pause
  exit /b 1
)

rem --- �g���q��������� .ps1 ��t�^�i��ݒ�΍�j ---
for %%Z in ("%SCRIPT%") do (
  set "EXT=%%~xZ"
)
if /I not "%EXT%"==".ps1" set "SCRIPT=%SCRIPT%.ps1"

rem --- ��΃p�X�������i�X�y�[�X/���{��/UNC ���艻�j ---
for %%I in ("%SCRIPT%")  do set "ABS_SCRIPT=%%~fI"
for %%I in ("%HOSTS%")   do set "ABS_HOSTS=%%~fI"
if exist "%PASSFILE%" (
  for %%I in ("%PASSFILE%") do set "ABS_PASSFILE=%%~fI"
) else (
  set "ABS_PASSFILE="
)

rem --- �O��t�@�C���m�F ---
if not exist "%ABS_SCRIPT%" ( echo [ERROR] �X�N���v�g��������܂���: "%ABS_SCRIPT%" & goto :fail )
if not exist "%ABS_HOSTS%"  ( echo [ERROR] hosts.txt ��������܂���: "%ABS_HOSTS%"  & goto :fail )

if defined ABS_PASSFILE (
  set "PWMSG=Password : file %ABS_PASSFILE%"
) else (
  set "PWMSG=Password : prompt"
)

rem --- �ǉ��I�v�V�����i�J��Ԃ��j ---
set "REPEAT_OPTS="
if "%REPEAT%"=="1" set "REPEAT_OPTS=%REPEAT_OPTS% -Repeat -IntervalMinutes %INTERVAL_MIN%"
if not "%DURATION_MIN%"=="0" set "REPEAT_OPTS=%REPEAT_OPTS% -DurationMinutes %DURATION_MIN%"

rem --- PowerShell ���s�t�@�C���̌��o�ipwsh �D��j ---
set "PS=pwsh.exe"
where pwsh.exe >nul 2>&1 || set "PS=powershell.exe"

echo * ���s�J�n: %date% %time%
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

echo --- ���s�R�}���h ---
if defined ABS_PASSFILE (
  echo "%PS%" -NoProfile -ExecutionPolicy Bypass -File "%ABS_SCRIPT%" -HostsFile "%ABS_HOSTS%" -Username "%USERNAME%" -PasswordFile "%ABS_PASSFILE%" %REPEAT_OPTS%
) else (
  echo "%PS%" -NoProfile -ExecutionPolicy Bypass -File "%ABS_SCRIPT%" -HostsFile "%ABS_HOSTS%" -Username "%USERNAME%" %REPEAT_OPTS%
)
echo --------------------

rem --- ���s�i-File �ɂ͕K�� .ps1 �̐�΃p�X��n���j ---
if defined ABS_PASSFILE (
  "%PS%" -NoProfile -ExecutionPolicy Bypass -File "%ABS_SCRIPT%" -HostsFile "%ABS_HOSTS%" -Username "%USERNAME%" -PasswordFile "%ABS_PASSFILE%" %REPEAT_OPTS%
) else (
  "%PS%" -NoProfile -ExecutionPolicy Bypass -File "%ABS_SCRIPT%" -HostsFile "%ABS_HOSTS%" -Username "%USERNAME%" %REPEAT_OPTS%
)

set "RC=%ERRORLEVEL%"
echo * �I���R�[�h: %RC%
if not "%RC%"=="0" echo [ERROR] PowerShell �X�N���v�g���ُ�I�����܂����B��̃��b�Z�[�W���m�F���Ă�������.
goto :end

:fail
set "RC=1"

:end
popd >nul 2>&1
if "%DEBUG_HOLD%"=="1" pause
exit /b %RC%
