@echo off
setlocal EnableExtensions
rem Builds dist\K-FRTB-ASA.xlsm from the template and the VBA modules of this
rem repository checkout with tools\Build-Workbook.ps1.
set "REPO_ROOT=%~dp0.."
set "SCRIPT=%~dp0Build-Workbook.ps1"
set "OUTPUT=%REPO_ROOT%\dist\K-FRTB-ASA.xlsm"

if not exist "%SCRIPT%" (
    echo ERROR: PowerShell script not found:
    call :echovar SCRIPT
    echo.
    pause
    exit /b 1
)

echo FRTB-SA workbook build
echo.
echo   P = production: src\modules only
echo   D = development: src\modules and tests\modules; runs the calculation
echo       and the regression test after the build
echo.
set "KIND="
set /p "KIND=Production or development? [P/D]: "
set "ARGS="
if /i "%KIND%"=="P" set "ARGS=-Production"
if /i "%KIND%"=="D" set "ARGS=-RunTests"
if not defined ARGS (
    echo CANCELLED: nothing was built.
    echo.
    pause
    exit /b 2
)

if exist "%OUTPUT%" (
    echo The workbook exists:
    call :echovar OUTPUT
    set "OVERWRITE="
    set /p "OVERWRITE=Overwrite it? [Y/N]: "
    call :overwrite
    if errorlevel 1 (
        echo CANCELLED: nothing was built.
        echo.
        pause
        exit /b 2
    )
    set "ARGS=%ARGS% -Force"
)

echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" %ARGS%
set "RC=%ERRORLEVEL%"

echo.
if "%RC%"=="0" (
    echo SUCCESS: workbook built in dist\.
) else (
    echo FAILED: the build returned exit code %RC%.
)
echo.
pause
exit /b %RC%

:overwrite
if /i "%OVERWRITE%"=="Y" exit /b 0
exit /b 1

:echovar
rem Prints the value of the named variable, indented. Delayed expansion
rem inserts the value after the line is parsed, so ampersands, pipes,
rem redirections and parentheses in a path are printed, never run.
setlocal EnableDelayedExpansion
echo(  !%~1!
endlocal
exit /b 0
