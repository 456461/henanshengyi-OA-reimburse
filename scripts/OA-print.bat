@echo off
REM ==========================================================
REM   OA reimbursement print tool
REM   Double click to run. Menu driven.
REM ==========================================================
setlocal
chcp 65001 >nul 2>&1

set "HERE=%~dp0"
set "PSFILE=%HERE%OA-print.ps1"

if not exist "%PSFILE%" (
    echo.
    echo [ERROR] OA-print.ps1 not found in:
    echo         %HERE%
    echo.
    goto :end
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%PSFILE%" %*
set "RC=%ERRORLEVEL%"

if not "%RC%"=="0" (
    echo.
    echo [FAILED] exit code: %RC%
    echo A log file was written to this folder.
    echo Please send that log file for diagnosis.
    echo.
)

:end
echo.
pause
endlocal
