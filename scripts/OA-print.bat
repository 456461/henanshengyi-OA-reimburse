@echo off
REM ==========================================================
REM   OA reimbursement print tool - SINGLE ENTRY
REM
REM   No args   -> interactive menu (double click to use)
REM   With args -> pass through to OA-print.ps1 (CLI mode)
REM
REM   CLI examples:
REM     OA-print.bat -Mode Step1 -DryRun
REM     OA-print.bat -Mode Step2 -DryRun
REM     OA-print.bat -Mode Step1
REM     OA-print.bat -Mode Step2
REM     OA-print.bat -Mode Group -Group 1266 -DryRun
REM     OA-print.bat -Mode Diagnose
REM     OA-print.bat -Diagnose
REM
REM   Note: -ExecutionPolicy Bypass is REQUIRED.
REM         Local policy blocks unsigned .ps1 scripts.
REM ==========================================================
setlocal
chcp 65001 >nul 2>&1

set "HERE=%~dp0"
set "PSFILE=%HERE%OA-print.ps1"
set "RC=0"

if not exist "%PSFILE%" (
    echo.
    echo [ERROR] OA-print.ps1 not found in:
    echo         %HERE%
    echo.
    set "RC=1"
    goto :done
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%PSFILE%" %*
set "RC=%ERRORLEVEL%"

:done
if not "%RC%"=="0" (
    echo.
    echo [FAILED] exit code: %RC%
    echo A log file was written to this folder.
    echo Please send that log file for diagnosis.
    echo.
)

REM Pause only when launched by double-click (no args), so CLI use is not blocked
if "%~1"=="" (
    echo Press any key to exit...
    pause >nul
)

endlocal & exit /b %RC%
