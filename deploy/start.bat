@echo off
REM sheetops one-click launcher. Double-click this file.
REM All messages come from deploy\setup.ps1 (Traditional Chinese).
chcp 65001 >nul
cd /d "%~dp0.."
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0setup.ps1" %*
if errorlevel 1 (
  echo.
  echo [!] Setup failed. See the message above.
  pause
)
