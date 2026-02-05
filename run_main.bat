@echo off
setlocal
pushd "%~dp0"

set "PYTHON_EXE="
if exist ".venv\Scripts\python.exe" set "PYTHON_EXE=.venv\Scripts\python.exe"
if not defined PYTHON_EXE set "PYTHON_EXE=python"

echo Running src\main.py with %PYTHON_EXE%
%PYTHON_EXE% src\main.py
set "EXIT_CODE=%ERRORLEVEL%"

if not "%EXIT_CODE%"=="0" (
  echo.
  echo src\main.py exited with code %EXIT_CODE%
)
echo.
pause
popd
