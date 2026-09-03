@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "APP_DIR=%SCRIPT_DIR%..\"
set "OUT_EXE=%APP_DIR%pm.exe"
set "ICON=%APP_DIR%assets\acies.ico"
set "CSC=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"

if not exist "%CSC%" (
    echo [ERROR] csc.exe compiler not found at: %CSC%
    exit /b 1
)

echo Compiling silent launcher pm.exe...
"%CSC%" /target:winexe /optimize+ /r:System.Windows.Forms.dll /win32icon:"%ICON%" /out:"%OUT_EXE%" "%SCRIPT_DIR%launcher.cs"
if %ERRORLEVEL% equ 0 (
    echo [SUCCESS] Compiled %OUT_EXE%
) else (
    echo [FAILED] Compilation failed.
)

exit /b %ERRORLEVEL%
