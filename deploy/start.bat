@echo off
REM sheetops one-click launcher. Double-click this file.
REM All messages come from deploy\setup.ps1 (Traditional Chinese).
chcp 65001 >nul
cd /d "%~dp0.."
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0setup.ps1" %*
echo.
echo ============================================
echo  Server stopped. Press any key to close.
echo ============================================
pause >nul
