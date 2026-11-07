@echo off
setlocal

set "HERE=%~dp0"

powershell -NoProfile -ExecutionPolicy Bypass -Command "$s = Get-ChildItem -LiteralPath $env:HERE -Filter '*.ps1' -ErrorAction SilentlyContinue | Select-Object -First 1; if (-not $s) { Write-Host 'ERROR: .ps1 not found next to this .bat' -ForegroundColor Red; exit 1 }; & $s.FullName %*"
set "RC=%ERRORLEVEL%"

echo.
if not "%RC%"=="0" echo Exit code: %RC%
pause
endlocal
exit /b %RC%
