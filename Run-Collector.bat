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

rem --- �����R�[�h�i���{��R���\�[���p / Shift-JIS�j---
chcp 932 >nul 2>&1

rem --- ���s�f�B���N�g���ֈړ��iUNC/OneDrive �ł� pushd �ň���j ---
set "BASE=%~dp0"
pushd "%BASE%" >nul 2>&1
if errorlevel 1 (
  echo [ERROR] ��ƃt�H���_�ֈړ��ł��܂���: "%BASE%"
  if "%DEBUG_HOLD%"=="1" pause
  exit /b 1
)

rem --- PowerShell ���s�t�@�C���̌��o�ipwsh �D��j ---
set "PS=pwsh.exe"
where pwsh.exe >nul 2>&1 || set "PS=powershell.exe"

rem --- �O��t�@�C���m�F ---
if not exist "%SCRIPT%"  ( echo [ERROR] �X�N���v�g��������܂���: "%CD%\%SCRIPT%"  & goto :fail )
if not exist "%HOSTS%"   ( echo [ERROR] hosts.txt ��������܂���: "%CD%\%HOSTS%"   & goto :fail )

if exist "%PASSFILE%" (
  set "PWMSG=Password : file %CD%\%PASSFILE%"
) else (
  set "PWMSG=Password : prompt"
)

rem --- PowerShell �����̑g�ݗ��āi\" �͎g��Ȃ� / �t���p�X�Ŋm���Ɂj ---
set "PSARGS=-NoProfile -ExecutionPolicy Bypass -File ""%CD%\%SCRIPT%"" -HostsFile ""%CD%\%HOSTS%"" -Username ""%USERNAME%"""
if exist "%PASSFILE%" set "PSARGS=%PSARGS% -PasswordFile ""%CD%\%PASSFILE%"""
if "%REPEAT%"=="1"     set "PSARGS=%PSARGS% -Repeat -IntervalMinutes %INTERVAL_MIN%"
if not "%DURATION_MIN%"=="0" set "PSARGS=%PSARGS% -DurationMinutes %DURATION_MIN%"

echo * ���s�J�n: %date% %time%
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

echo --- ���s�R�}���h ---
echo %PS% %PSARGS%
echo --------------------

"%PS%" %PSARGS%
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
