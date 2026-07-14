@echo off
setlocal
pushd "%~dp0"

node -e "require.resolve('exceljs'); require.resolve('json5'); require.resolve('@msgpack/msgpack')" >nul 2>&1
if errorlevel 1 call npm ci --no-audit --no-fund
if errorlevel 1 goto :failed

node src\cli.js --target client
if errorlevel 1 goto :failed

popd
exit /b 0

:failed
set "EXIT_CODE=%ERRORLEVEL%"
echo.
echo Export client failed with exit code %EXIT_CODE%.
pause
popd
exit /b %EXIT_CODE%
