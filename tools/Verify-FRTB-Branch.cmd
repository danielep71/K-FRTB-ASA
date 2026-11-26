@echo off
setlocal EnableExtensions
rem Verifies a branch of origin with the unattended Excel check:
rem tools\Check-FRTB-Excel.ps1 checks the branch out, runs the Excel checks and
rem writes a zipped bundle under test-results\. The branch is the first argument
rem and quick, or left out for the full check, the second. Double-clicked, it asks
rem for both. A quick check is not evidence; a full one takes about ten minutes.
set "REPO_ROOT=%~dp0.."
cd /d "%REPO_ROOT%"
set "SCRIPT=%~dp0Check-FRTB-Excel.ps1"
set "DEFAULT_BRANCH=release/v0.3.0"

if not exist "%SCRIPT%" (
    echo ERROR: PowerShell script not found:
    call :echovar SCRIPT
    echo.
    pause
    exit /b 1
)

set "BRANCH=%~1"
set "MODE=%~2"
if not defined BRANCH (
    set /p "BRANCH=Branch to verify, as on GitHub [%DEFAULT_BRANCH%]: "
    set /p "MODE=Quick check, not evidence, or full check? [Q/F]: "
)
if not defined BRANCH set "BRANCH=%DEFAULT_BRANCH%"
if not defined MODE set "MODE=F"
set "BRANCH=%BRANCH:"=%"
set "MODE=%MODE:"=%"
if not defined BRANCH (
    echo ERROR: no branch given.
    echo.
    pause
    exit /b 2
)

set "QUICK="
set "KNOWN="
if /i "%MODE%"=="Q" set "QUICK=-Quick"
if /i "%MODE%"=="quick" set "QUICK=-Quick"
if /i "%MODE%"=="-quick" set "QUICK=-Quick"
if /i "%MODE%"=="/quick" set "QUICK=-Quick"
if defined QUICK set "KNOWN=1"
if /i "%MODE%"=="F" set "KNOWN=1"
if /i "%MODE%"=="full" set "KNOWN=1"
if not defined KNOWN (
    echo ERROR: the second argument must be quick, or left out for the full check:
    call :echovar MODE
    echo.
    pause
    exit /b 2
)

echo K-FRTB-ASA Excel check of a branch
echo.
echo Branch:
call :echovar BRANCH
echo Repository:
call :echovar REPO_ROOT
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -Branch "%BRANCH%" %QUICK%
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

:echovar
rem Prints the value of the named variable, indented. Delayed expansion
rem inserts the value after the line is parsed, so ampersands, pipes,
rem redirections and parentheses in a path are printed, never run.
setlocal EnableDelayedExpansion
echo(  !%~1!
endlocal
exit /b 0
