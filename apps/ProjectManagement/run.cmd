@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "VENV_PYTHON=%SCRIPT_DIR%.venv\Scripts\python.exe"

if exist "%VENV_PYTHON%" (
    "%VENV_PYTHON%" "%SCRIPT_DIR%main.py" %*
) else (
    python "%SCRIPT_DIR%main.py" %*
)

exit /b %ERRORLEVEL%
