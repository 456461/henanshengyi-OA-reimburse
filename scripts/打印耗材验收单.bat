@echo off
setlocal enabledelayedexpansion

set "HERE=%~dp0"

rem try switch console to UTF-8 to avoid mojibake
chcp 65001 >nul 2>&1

powershell -NoProfile -ExecutionPolicy Bypass -Command "$s = Get-ChildItem -LiteralPath $env:HERE -Filter '*.ps1' -ErrorAction SilentlyContinue | Select-Object -First 1; if (-not $s) { Write-Host 'ERROR: .ps1 not found next to this .bat' -ForegroundColor Red; exit 1 }; & $s.FullName %*"
set "RC=%ERRORLEVEL%"

echo.
if not "%RC%"=="0" (
    echo [FAILED] exit code: %RC%
    echo.
    echo A log file was written to this folder:
    echo     "*.txt"  ^(name starts with the characters for print-log^)
    echo.
    echo Please send that log file to support for diagnosis.
) else (
    echo [OK] all done.
)
echo.
pause
endlocal
exit /b %RC%
