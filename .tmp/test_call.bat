@echo off
setlocal
call :try_base "py -3.10"
echo done
exit /b 0
:try_base
set "CANDIDATE=%~1"
call %CANDIDATE% -c "import sys" >nul 2>&1
exit /b 0
