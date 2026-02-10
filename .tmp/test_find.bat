@echo off
setlocal EnableExtensions EnableDelayedExpansion
call :find_base_python
echo BASE=%BASE_PY_CMD%
exit /b 0

:find_base_python
set "BASE_PY_CMD="
for /f "delims=" %%P in ('where python 2^>nul') do (
  call :check_python_exe "%%P"
  if defined BASE_PY_CMD exit /b 0
)
where py >nul 2>&1
if "%ERRORLEVEL%"=="0" (
  call :check_py_launcher "3.12"
  if defined BASE_PY_CMD exit /b 0
  call :check_py_launcher "3.11"
  if defined BASE_PY_CMD exit /b 0
  call :check_py_launcher "3.10"
  if defined BASE_PY_CMD exit /b 0
  call :check_py_launcher "3"
  if defined BASE_PY_CMD exit /b 0
)
exit /b 0

:check_python_exe
set "CANDIDATE=%~1"
"%CANDIDATE%" -c "import sys, venv; exit(0 if sys.version_info[:2]>=(3,10) else 1)" >nul 2>&1
if "%ERRORLEVEL%"=="0" (
  set "BASE_PY_CMD=\"%CANDIDATE%\""
)
exit /b 0

:check_py_launcher
set "PYVER=%~1"
py -%PYVER% -c "import sys, venv; exit(0 if sys.version_info[:2]>=(3,10) else 1)" >nul 2>&1
if "%ERRORLEVEL%"=="0" (
  set "BASE_PY_CMD=py -%PYVER%"
)
exit /b 0
