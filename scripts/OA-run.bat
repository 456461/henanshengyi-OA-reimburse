@echo off
REM ==========================================================
REM   OA reimbursement print - command line mode
REM
REM   Examples:
REM     OA-run.bat -Mode Step1 -DryRun     preview docx
REM     OA-run.bat -Mode Step2 -DryRun     preview pdf
REM     OA-run.bat -Mode Step1             print docx
REM     OA-run.bat -Mode Step2             print pdf
REM     OA-run.bat -Mode Group -Group 1266 -DryRun
REM     OA-run.bat -Mode Diagnose          diagnose env
REM
REM   Note: double click OA-print.bat for the menu instead.
REM ==========================================================
setlocal
chcp 65001 >nul 2>&1

set "HERE=%~dp0"
set "PSFILE=%HERE%OA-print.ps1"

if not exist "%PSFILE%" (
    echo [ERROR] OA-print.ps1 not found in %HERE%
    goto :end
)

REM -ExecutionPolicy Bypass is required: local policy blocks unsigned scripts
powershell -NoProfile -ExecutionPolicy Bypass -File "%PSFILE%" %*
set "RC=%ERRORLEVEL%"

:end
echo.
pause
endlocal
