@echo off
setlocal EnableExtensions EnableDelayedExpansion

rem ===================== �ݒ� =====================
set "BASE=%~dp0"
set "SCRIPT=%BASE%Get-CiscoInterfaces-PerIF.ps1"
set "HOSTS=%BASE%hosts.txt"
set "PASSFILE=%BASE%password.txt"
set "USERNAME=cisco"

rem �J��Ԃ�: 1=�L��, 0=1��̂�
set "REPEAT=1"
set "INTERVAL_MIN=60"
rem �����s����(��)�B0=�������iREPEAT=1�̎��̂ݗL���j
set "DURATION_MIN=0"

rem ���s���ɉ�ʂ���Ȃ�: 1=�L��/0=����
set "DEBUG_HOLD=1"
rem =================================================

rem --- PowerShell ���s�t�@�C���̌��o�ipwsh �D��j ---
set "PS=pwsh.exe"
where pwsh.exe >nul 2>&1 || set "PS=powershell.exe"

rem --- �O��t�@�C���m�F ---
if not exist "%SCRIPT%" (
  echo [ERROR] �X�N���v�g��������܂���: "%SCRIPT%"
  goto :fail
)
if not exist "%HOSTS%" (
  echo [ERROR] hosts.txt ��������܂���: "%HOSTS%"
  goto :fail
)

rem �\���p���b�Z�[�W
if exist "%PASSFILE%" (
  set "PWMSG=Password : file %PASSFILE%"
) else (
  set "PWMSG=Password : prompt"
)

rem --- PowerShell �����𐳂����g�ݗ��āi\" �͎g��Ȃ��j ---
set "PSARGS=-NoProfile -ExecutionPolicy Bypass -File ""%SCRIPT%"" -HostsFile ""%HOSTS%"" -Username ""%USERNAME%"""
if exist "%PASSFILE%" set "PSARGS=%PSARGS% -PasswordFile ""%PASSFILE%"""
if "%REPEAT%"=="1" set "PSARGS=%PSARGS% -Repeat -IntervalMinutes %INTERVAL_MIN%"
if not "%DURATION_MIN%"=="0" set "PSARGS=%PSARGS% -DurationMinutes %DURATION_MIN%"

echo * ���s�J�n: %date% %time%
echo   PS        : %PS%
echo   Script    : %SCRIPT%
echo   Hosts     : %HOSTS%
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
if "%DEBUG_HOLD%"=="1" pause
exit /b %RC%
