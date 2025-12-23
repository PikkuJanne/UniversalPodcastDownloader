@echo off
setlocal

REM Folder where this .bat lives
set "SCRIPT_DIR=%~dp0"

REM Run the PowerShell script with relaxed execution policy
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%UniversalPodcastDownloader.ps1"

echo.
echo Done. You can close this window.
pause
endlocal
