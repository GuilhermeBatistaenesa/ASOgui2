@echo off
setlocal
pushd "%~dp0"

set "ARG=%~1"
set "PS_SCRIPT=scripts\build_aso_zip.ps1"

if "%ARG%"=="" (
  powershell -ExecutionPolicy Bypass -File "%PS_SCRIPT%"
) else (
  if /I "%ARG%"=="major" (
    powershell -ExecutionPolicy Bypass -File "%PS_SCRIPT%" -Bump major
  ) else if /I "%ARG%"=="minor" (
    powershell -ExecutionPolicy Bypass -File "%PS_SCRIPT%" -Bump minor
  ) else if /I "%ARG%"=="patch" (
    powershell -ExecutionPolicy Bypass -File "%PS_SCRIPT%" -Bump patch
  ) else (
    powershell -ExecutionPolicy Bypass -File "%PS_SCRIPT%" -Version "%ARG%"
  )
)

set "EXIT_CODE=%ERRORLEVEL%"
if not "%EXIT_CODE%"=="0" (
  echo.
  echo build_aso_zip.ps1 failed with code %EXIT_CODE%
)
echo.
pause
popd
