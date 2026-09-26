@echo off
rem Double-click to launch the GUI (no console window).
rem Keep this file ASCII-only: cmd.exe reads .cmd in the OEM codepage,
rem so UTF-8 Chinese here would come out garbled.
setlocal
set "SCRIPT=%~dp0export-gui.ps1"
if not exist "%SCRIPT%" (
  echo Cannot find export-gui.ps1 -- is the scripts folder complete?
  pause
  exit /b 1
)
start "" "%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" ^
  -NoProfile -STA -WindowStyle Hidden -ExecutionPolicy Bypass -File "%SCRIPT%" %*
endlocal
