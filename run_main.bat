@echo off
setlocal EnableDelayedExpansion
title ASOgui - Run Main
pushd "%~dp0"

set "PYTHON_CMD="
set "BASE_PY_CMD="
set "DIAG=%RUN_MAIN_DIAG%"

if defined DIAG (
  echo ===== RUN_MAIN_DIAG =====
  echo Current dir: %CD%
  where py >nul 2>&1 && (echo py launcher found) || (echo py launcher NOT found)
  where python >nul 2>&1 && (echo python found in PATH) || (echo python NOT found in PATH)
  if exist ".venv\pyvenv.cfg" (
    echo --- .venv\pyvenv.cfg ---
    type ".venv\pyvenv.cfg"
    echo -----------------------
  ) else (
    echo .venv\pyvenv.cfg not found
  )
  echo =========================
)

if exist ".venv\Scripts\python.exe" set "PYTHON_CMD=.venv\Scripts\python.exe"
if not defined PYTHON_CMD (
  where py >nul 2>&1 && set "PYTHON_CMD=py -3.10"
)
if not defined PYTHON_CMD (
  where py >nul 2>&1 && set "PYTHON_CMD=py -3.11"
)
if not defined PYTHON_CMD (
  where python >nul 2>&1 && set "PYTHON_CMD=python"
)
set "BASE_PY_CMD=%PYTHON_CMD%"
if not defined PYTHON_CMD (
  echo Python nao encontrado. Instale Python 3.10+.
  echo.
  set "EXIT_CODE=1"
  goto :end
)

if not exist ".venv\Scripts\python.exe" (
  echo Criando .venv...
  if defined DIAG (
    echo Usando: %PYTHON_CMD% -m venv .venv
  )
  call %PYTHON_CMD% -m venv .venv
  if not exist ".venv\Scripts\python.exe" (
    echo Falha ao criar .venv.
    echo.
    set "EXIT_CODE=1"
    goto :end
  )
  set "PYTHON_CMD=.venv\Scripts\python.exe"
)

call %PYTHON_CMD% -c "import sys" >nul 2>&1
if not "%ERRORLEVEL%"=="0" (
  echo Venv quebrada. Recriando .venv...
  if exist ".venv" (
    if not exist ".venv\_delete.lock" (
      echo.> ".venv\_delete.lock"
    )
    cmd /c "rmdir /s /q .venv" >nul 2>&1
  )
  set "PYTHON_CMD=%BASE_PY_CMD%"
  call %PYTHON_CMD% -m venv .venv
  if not exist ".venv\Scripts\python.exe" (
    echo Falha ao criar .venv.
    echo.
    set "EXIT_CODE=1"
    goto :end
  )
  set "PYTHON_CMD=.venv\Scripts\python.exe"
)

echo Running src\main.py with %PYTHON_CMD%
if defined DIAG (
  call %PYTHON_CMD% -c "import sys; print('Python:', sys.version); print('Executable:', sys.executable)"
)
call %PYTHON_CMD% -c "import openpyxl" >nul 2>&1
if not "%ERRORLEVEL%"=="0" (
  echo Installing dependencies from requirements.txt...
  if defined DIAG (
    echo Pip command: %PYTHON_CMD% -m pip install -r requirements.txt
  )
  call %PYTHON_CMD% -m pip install -r requirements.txt
)
call %PYTHON_CMD% src\main.py
set "EXIT_CODE=%ERRORLEVEL%"

if not "%EXIT_CODE%"=="0" (
  echo.
  echo src\main.py exited with code %EXIT_CODE%
)
:end
echo.
pause
popd
exit /b %EXIT_CODE%
