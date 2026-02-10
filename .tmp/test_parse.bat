@echo off
setlocal EnableDelayedExpansion
set "VENV_DIR=%LOCALAPPDATA%\ASOgui\.venv"
if exist "%VENV_DIR%\pyvenv.cfg" (
  for /f "tokens=2 delims==" %%A in ('findstr /b /c:"home =" "%VENV_DIR%\pyvenv.cfg"') do set "VENV_HOME=%%A"
  if defined VENV_HOME (
    set "VENV_HOME=!VENV_HOME:~1!"
    if not exist "!VENV_HOME!" (
      echo Venv base python nao encontrado: !VENV_HOME!
    )
  )
)
exit /b 0
