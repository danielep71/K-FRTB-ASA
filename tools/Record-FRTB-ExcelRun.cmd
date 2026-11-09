@echo off
setlocal EnableExtensions
rem Double-click to record an Excel run of the checked-out commit:
rem tools\Record-FRTB-ExcelRun.ps1 builds, compiles and compares both
rem workbooks, writes the evidence bundle under test-results\ and validates it.
set "REPO_ROOT=%~dp0.."
cd /d "%REPO_ROOT%"
set "SCRIPT=%~dp0Record-FRTB-ExcelRun.ps1"
if not exist "%SCRIPT%" (
    echo ERROR: PowerShell script not found:
    echo   %SCRIPT%
    pause
    exit /b 1
)
echo K-FRTB-SA Excel run
echo Repository: %CD%
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%"
set "RC=%ERRORLEVEL%"
echo.
if "%RC%"=="0" (
    echo RECORDED: every stage passed and the record is valid.
) else if "%RC%"=="1" (
    echo RECORDED WITH FINDINGS: see the stage list and the validator outcome above.
) else (
    echo NOT RECORDED: the run stopped with exit code %RC%; see the message above.
)
echo.
pause
exit /b %RC%
