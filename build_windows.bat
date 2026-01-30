@echo off
setlocal
pushd "%~dp0"

powershell -ExecutionPolicy Bypass -File "scripts\build_windows.ps1"
set "EXIT_CODE=%ERRORLEVEL%"

if not "%EXIT_CODE%"=="0" (
  echo.
  echo build_windows.ps1 failed with code %EXIT_CODE%
)
echo.
pause
popd
