@echo off
setlocal EnableExtensions DisableDelayedExpansion

REM The script owns the argument-free guided pause without reparsing arguments.
set "UPD_LAUNCHER=1"
set "UPD_LAUNCHER_SCRIPT=%~dp0UniversalPodcastDownloader.ps1"
if not exist "%UPD_LAUNCHER_SCRIPT%" goto missing_script

REM Remove an inherited shadow so ERRORLEVEL reports the actual child status.
set "ERRORLEVEL="
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%UPD_LAUNCHER_SCRIPT%" %*
set "UPD_LAUNCHER_EXIT=%ERRORLEVEL%"
endlocal & exit /b %UPD_LAUNCHER_EXIT%

:missing_script
>&2 echo ERROR: UniversalPodcastDownloader.ps1 is missing next to the launcher.
endlocal & exit /b 1
