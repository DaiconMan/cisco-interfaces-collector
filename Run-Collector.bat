@echo off
setlocal ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION

rem ===================== �ݒ�i�K�v�ɉ����ĕҏW�j =====================
set "BASE=%~dp0"
set "SCRIPT=%BASE%Get-CiscoInterfaces-PerIF.ps1"
set "HOSTS=%BASE%hosts.txt"
set "PASSFILE=%BASE%password.txt"
set "USERNAME=cisco

rem �J��Ԃ��ݒ�FREPEAT=1 �ŗL���A0 ��1��̂�
set "REPEAT=1"
set "INTERVAL_MIN=60"
rem �C�ӁF�����s���ԁi���j�B0 �Ȃ疳����
set "DURATION_MIN=0"

rem ���s���ɉ�ʂ���Ȃ��i1=�L��/0=�����j
set "DEBUG_HOLD=1"
rem ====================================================================

rem --- PowerShell ���s�t�@�C���̌��o�ipwsh �D��j ---
set "PS=pwsh.exe"
where "%PS%" >nul 2>&1
if errorlevel 1 set "PS=powershell.exe"

rem --- �O��t�@�C���m�F ---
if not exist "%SCRIPT%" (
  echo [ERROR] �X�N���v�g��������܂���: "%SCRIPT%"
  if "%DEBUG_HOLD%"=="1" pause
  exit /b 1
)
if not exist "%HOSTS%" (
  echo [ERROR] hosts.txt ��������܂���: "%HOSTS%"
  if "%DEBUG_HOLD%"=="1" pause
  exit /b 1
)

rem --- �I�v�V�����̑g�ݗ��āi���p���͑f���Ɂj ---
set "PWOPT="
if exist "%PASSFILE%" set "PWOPT=-PasswordFile \"%PASSFILE%\""

set "REPEATOPT="
if "%REPEAT%"=="1" (
  set "REPEATOPT=-Repeat -IntervalMinutes %INTERVAL_MIN%"
  if not "%DURATION_MIN%"=="0" set "REPEATOPT=%REPEATOPT% -DurationMinutes %DURATION_MIN%"
)

echo * ���s�J�n: %DATE% %TIME%
echo   PS        : %PS%
echo   Script    : %SCRIPT%
echo   Hosts     : %HOSTS%
echo   Username  : %USERNAME%
if defined PWOPT (echo   Password : file %PASSFILE%) else (echo   Password : prompt)
if defined REPEATOPT (echo   Repeat   : every %INTERVAL_MIN% min ^| Duration=%DURATION_MIN% min) else (echo   Repeat   : no)

"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -HostsFile "%HOSTS%" -Username "%USERNAME%" %PWOPT% %REPEATOPT%
set "RC=%ERRORLEVEL%"

echo * �I���R�[�h: %RC%
if not "%RC%"=="0" (
  echo [ERROR] PowerShell �X�N���v�g���ُ�I�����܂����B��̃��b�Z�[�W���m�F���Ă��������B
  if "%DEBUG_HOLD%"=="1" pause
)
endlocal & exit /b %RC%
