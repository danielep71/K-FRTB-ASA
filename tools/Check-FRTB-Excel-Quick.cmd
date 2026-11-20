@echo off
setlocal EnableExtensions
rem Double-click to run the quick Excel check of a branch, unattended: one
rem development build with the regression test, then the acceptance checks.
rem It is not evidence: record a full run with Check-FRTB-Excel.cmd before an
rem integration. Default branch: release/v0.3.0; pass another as the first argument.
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
if not defined BRANCH set "BRANCH=release/v0.3.0"
echo K-FRTB-ASA quick Excel check of %BRANCH%
echo Repository: %CD%
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -Branch "%BRANCH%" -Quick
set "RC=%ERRORLEVEL%"
echo.
if "%RC%"=="0" (
    echo PASS: the quick checks passed. Send the zip shown above.
) else if "%RC%"=="1" (
    echo FINDINGS: some checks failed; see the summary above. Send the zip shown above.
) else (
    echo NOT STARTED: the check stopped with exit code %RC%; see the message above.
)
echo.
pause
exit /b %RC%
