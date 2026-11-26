@echo off
setlocal EnableExtensions
rem Synchronizes the VBA modules of this repository checkout into an FRTB-SA
rem workbook. Drag the .xlsm onto this file, or pass its path and, as a second
rem argument, P or D for a production or a development workbook; without the
rem second argument it asks. Without a workbook, it builds a development
rem workbook from the template in test-results\, replacing the last one.
rem The directory is not changed: a relative workbook path is the caller's.
for %%I in ("%~dp0..") do set "REPO_ROOT=%%~fI"
set "SCRIPT=%~dp0Sync-FRTB-VBA.ps1"
set "BUILD_SCRIPT=%~dp0Build-Workbook.ps1"
set "FRESH=%REPO_ROOT%\test-results\K-FRTB-ASA-Development.xlsm"

if not exist "%SCRIPT%" (
    echo ERROR: PowerShell script not found:
    call :echovar SCRIPT
    echo.
    pause
    exit /b 1
)

set "WORKBOOK=%~1"
if not defined WORKBOOK goto :fresh
if not exist "%WORKBOOK%" (
    echo ERROR: Workbook not found:
    call :echovar WORKBOOK
    echo.
    pause
    exit /b 1
)

set "MODE=%~2"
if not defined MODE (
    echo   P = production: src\modules only
    echo   D = development: src\modules and tests\modules
    echo.
    set /p "MODE=Production or development workbook? [P/D]: "
)
if defined MODE set "MODE=%MODE:"=%"
set "MODE_ARG="
if /i "%MODE%"=="P" set "MODE_ARG=-Mode Production"
if /i "%MODE%"=="Production" set "MODE_ARG=-Mode Production"
if /i "%MODE%"=="D" set "MODE_ARG=-Mode Development"
if /i "%MODE%"=="Development" set "MODE_ARG=-Mode Development"
if not defined MODE_ARG (
    echo CANCELLED: the mode must be P or D; nothing was changed.
    echo.
    pause
    exit /b 2
)

echo FRTB-SA VBA synchronization from the repository checkout
echo.
echo Repository:
call :echovar REPO_ROOT
echo Workbook:
call :echovar WORKBOOK
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -WorkbookPath "%WORKBOOK%" -RepoRoot "%REPO_ROOT%" %MODE_ARG%
set "RC=%ERRORLEVEL%"

echo.
if "%RC%"=="0" (
    echo SUCCESS: VBA synchronized.
) else if "%RC%"=="2" (
    echo CANCELLED: the workbook was not changed.
) else (
    echo FAILED: the synchronization returned exit code %RC%.
)
echo.
pause
exit /b %RC%

:fresh
echo FRTB-SA development workbook from the template
echo.
echo No workbook given. A development workbook is built from the template,
echo with the modules of src\modules and tests\modules, replacing the last one:
call :echovar FRESH
echo.

if not exist "%BUILD_SCRIPT%" (
    echo ERROR: PowerShell script not found:
    call :echovar BUILD_SCRIPT
    echo.
    pause
    exit /b 1
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%BUILD_SCRIPT%" -OutputPath "%FRESH%" -Force
set "RC=%ERRORLEVEL%"

echo.
if "%RC%"=="0" (
    echo SUCCESS: development workbook built.
    echo Open it, run Debug - Compile VBAProject, then the regression test:
    call :echovar FRESH
) else (
    echo FAILED: the build returned exit code %RC%.
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
