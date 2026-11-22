@echo off
rem ============================================================
rem  DeepSeek Key Manager - development launcher
rem  Runs the main program with the system Python.
rem ============================================================
cd /d "%~dp0"

if not exist "DeepSeek_Key_Manager.py" (
    echo [ERROR] DeepSeek_Key_Manager.py not found in %CD%
    pause
    exit /b 1
)

python DeepSeek_Key_Manager.py

if errorlevel 1 (
    echo.
    echo [WARN] The program exited with an error. Press any key to close.
    pause >nul
)
