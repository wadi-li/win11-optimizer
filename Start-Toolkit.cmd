@echo off
rem Win11 Boot/RAM/Battery Toolkit launcher: Windows PowerShell 5.1, ExecutionPolicy Bypass for this process only
rem No arguments: web panel in the browser (Start-Toolkit.cmd -Mode Menu = console menu)
setlocal
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
set "SCRIPT=%~dp0Win11-BootRamBattery-Toolkit.ps1"
if not exist "%SCRIPT%" (echo Win11-BootRamBattery-Toolkit.ps1 not found next to this file. Unpack Win11Toolkit.zip completely. & pause & exit /b 1)
if "%~1"=="" (
  "%PS%" -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -Mode Panel
) else (
  "%PS%" -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" %*
)
set "RC=%ERRORLEVEL%"
if "%~1"=="" if not "%RC%"=="0" pause
exit /b %RC%
