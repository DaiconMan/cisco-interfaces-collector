@echo off
setlocal EnableExtensions EnableDelayedExpansion

rem ===================== �ݒ�i�K�v�ɉ����ĕύX�j =====================
rem �� ���� .bat �� .ps1 / hosts.txt / password.txt �͓����t�H���_�ɒu���O��
set "SCRIPT=Get-CiscoInterfaces-PerIF.ps1"
set "HOSTS=hosts.txt"
set "PASSFILE=password.txt"
set "USERNAME=cisco"

rem �J��Ԃ��ݒ�
set "REPEAT=1"
set "INTERVAL_MIN=60"
set "DURATION_MIN=0"   rem 0=�������iREPEAT=1�̂Ƃ������L���j

rem ���s���ɉ�ʂ���Ȃ��i1=�L��/0=�����j
set "DEBUG_HOLD=1"
rem =====================================================================

rem --- ���s�f�B���N�g���ֈړ��iUNC�Ȃ�ꎞ�h���C�u�����j ---
set "BASE=%~dp0"
pushd "%BASE%" >nul 2>&1
if errorlevel 1 (
  echo [ERROR] ��ƃt�H���_�ֈړ��ł��܂���: "%BASE%"
  if "%DEBUG_HOLD%"=="1" pause
  exit /b 1
)

rem --- �R���\�[����UTF-8�Ɂi�\������p�B���s���Ă������j ---
chcp 65001 >nul 2>&1

rem --- PowerShell ���s�t�@�C���̌��o�ipwsh �D��j ---
set "PS=pwsh.exe"
where pwsh.exe >nul 2>&1 || set "PS=powershell.exe"

rem --- �O��t�@�C���m�F�i���΃p�X��OK�j ---
if not exist ".\%SCRIPT%"  ( echo [ERROR] �X�N���v�g��������܂���: ".\%SCRIPT%"  & goto :fail )
if not exist ".\%HOSTS%"   ( echo [ERROR] hosts.txt ��������܂���: ".\%HOSTS%"   & goto :fail )
set "PWMSG=Password : prompt"
if exist ".\%PASSFILE%" set "PWMSG=Password : file .\%PASSFILE%"

rem --- PowerShell �����𑊑΃p�X�őg�ݗ��āi\" �͎g��Ȃ��j ---
set "PSARGS=-NoProfile -ExecutionPolicy Bypass -File "".\%SCRIPT%"" -HostsFile "".\%HOSTS%"" -Username ""%USERNAME%"""
if exist ".\%PASSFILE%" set "PSARGS=%PSARGS% -PasswordFile "".\%PASSFILE%"""
if "%REPEAT%"=="1"       set "PSARGS=%PSARGS% -Repeat -IntervalMinutes %INTERVAL_MIN%"
if not "%DURATION_MIN%"=="0" set "PSARGS=%PSARGS% -DurationMinutes %DURATION_MIN%"

echo * ���s�J�n: %date% %time%
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
echo * �I���R�[�h: %RC%
if not "%RC%"=="0" echo [ERROR] PowerShell �X�N���v�g���ُ�I�����܂����B��̃��b�Z�[�W���m�F���Ă�������.

goto :end

:fail
set "RC=1"

:end
popd >nul 2>&1
if "%DEBUG_HOLD%"=="1" pause
exit /b %RC%
