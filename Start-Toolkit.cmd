@echo off
rem Win11 Boot/RAM/Battery Toolkit launcher: Windows PowerShell 5.1, ExecutionPolicy Bypass for this process only
setlocal
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0Win11-BootRamBattery-Toolkit.ps1" %*
set "RC=%ERRORLEVEL%"
if "%~1"=="" if not "%RC%"=="0" pause
exit /b %RC%
