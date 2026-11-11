@echo off
setlocal EnableExtensions
rem Double-click to run every Excel check of a branch, unattended:
rem tools\Check-FRTB-Excel.ps1 checks the branch out, records the Excel run,
rem runs the acceptance, build, Trust Center and synchronization checks and
rem writes a zipped bundle under test-results\. Default branch:
rem release/v0.2.0; pass another as the first argument.
set "REPO_ROOT=%~dp0.."
cd /d "%REPO_ROOT%"
set "SCRIPT=%~dp0Check-FRTB-Excel.ps1"
if not exist "%SCRIPT%" (
    echo ERROR: PowerShell script not found:
    echo   %SCRIPT%
    pause
    exit /b 1
)
set "BRANCH=%~1"
if not defined BRANCH set "BRANCH=release/v0.2.0"
echo K-FRTB-ASA Excel check of %BRANCH%
echo Repository: %CD%
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -Branch "%BRANCH%"
set "RC=%ERRORLEVEL%"
echo.
if "%RC%"=="0" (
    echo PASS: every check passed. Send the zip shown above.
) else if "%RC%"=="1" (
    echo FINDINGS: some checks failed; see the summary above. Send the zip shown above.
) else (
    echo NOT STARTED: the check stopped with exit code %RC%; see the message above.
)
echo.
pause
exit /b %RC%
