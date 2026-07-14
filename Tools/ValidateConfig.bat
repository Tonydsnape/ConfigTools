@echo off
setlocal
chcp 65001 >nul
pushd "%~dp0.."

node -e "require.resolve('exceljs'); require.resolve('json5'); require.resolve('@msgpack/msgpack')" >nul 2>&1
if errorlevel 1 call npm ci --no-audit --no-fund
if errorlevel 1 goto :tool_failed

node Tools\validate-config.js --input "." --report "Tools\ValidationReport.txt"
set "EXIT_CODE=%ERRORLEVEL%"
echo.
pause
popd
exit /b %EXIT_CODE%

:tool_failed
set "EXIT_CODE=%ERRORLEVEL%"
echo.
echo Validation tool failed with exit code %EXIT_CODE%.
pause
popd
exit /b %EXIT_CODE%
