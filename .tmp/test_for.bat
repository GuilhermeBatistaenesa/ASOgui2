@echo off
setlocal
set "VENV_DIR=%LOCALAPPDATA%\ASOgui\.venv"
for %%I in ("%VENV_DIR%") do set "VENV_DIR=%%~fI"
echo %VENV_DIR%
exit /b 0
