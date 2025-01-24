@echo off
setlocal enabledelayedexpansion

if not exist logs mkdir logs
set "LOG_FILE=logs\startup_%date:~-4,4%%date:~-7,2%%date:~-10,2%_%time:~0,2%%time:~3,2%%time:~6,2%.log"

echo [%date% %time%] ========== Starting Audio Switcher Setup ========== >> "%LOG_FILE%"
echo [%date% %time%] Current directory: %CD% >> "%LOG_FILE%"
echo [%date% %time%] Python version: >> "%LOG_FILE%"
python --version >> "%LOG_FILE%" 2>&1

if not defined IS_ADMIN (
    set IS_ADMIN=1
    NET SESSION >nul 2>&1
    if %ERRORLEVEL% neq 0 (
        echo Requesting admin rights...
        powershell -Command "Start-Process '%~dpnx0' -Verb RunAs"
        exit /b
    )
)

echo [%date% %time%] Setting up Python environment... >> "%LOG_FILE%"
set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%" >> "%LOG_FILE%" 2>&1
set "VENV_DIR=%SCRIPT_DIR%venv"
set "PYTHON_DIR=%VENV_DIR%\Scripts"

echo [%date% %time%] Checking virtual environment... >> "%LOG_FILE%"
if not exist "%VENV_DIR%" (
    echo [%date% %time%] Creating new virtual environment... >> "%LOG_FILE%"
    python -m venv "%VENV_DIR%" >> "%LOG_FILE%" 2>&1
    if !ERRORLEVEL! neq 0 (
        echo [%date% %time%] ERROR: Failed to create virtual environment >> "%LOG_FILE%"
        echo Failed to create virtual environment. Check the logs at: %LOG_FILE%
        pause
        exit /b 1
    )
)

echo [%date% %time%] Activating virtual environment... >> "%LOG_FILE%"
call "%PYTHON_DIR%\activate.bat" >> "%LOG_FILE%" 2>&1
if %ERRORLEVEL% neq 0 (
    echo [%date% %time%] ERROR: Failed to activate virtual environment >> "%LOG_FILE%"
    echo Failed to activate virtual environment. Check the logs at: %LOG_FILE%
    pause
    exit /b 1
)

echo [%date% %time%] Checking Python in virtual environment: >> "%LOG_FILE%"
"%PYTHON_DIR%\python.exe" --version >> "%LOG_FILE%" 2>&1

echo [%date% %time%] Installing/Updating pip... >> "%LOG_FILE%"
"%PYTHON_DIR%\python.exe" -m pip install --upgrade pip >> "%LOG_FILE%" 2>&1

echo [%date% %time%] Installing requirements... >> "%LOG_FILE%"
"%PYTHON_DIR%\python.exe" -m pip install -r requirements.txt >> "%LOG_FILE%" 2>&1

echo [%date% %time%] Verifying icon.png exists... >> "%LOG_FILE%"
if not exist "icon.png" (
    echo [%date% %time%] Creating icon.png... >> "%LOG_FILE%"
    python -c "from icon import create_default_icon; create_default_icon()" >> "%LOG_FILE%" 2>&1
    if not exist "icon.png" (
        echo ERROR: Failed to create icon.png >> "%LOG_FILE%"
        echo Failed to create icon.png. Please check the logs.
        pause
        exit /b 1
    )
)

if not exist "config.json" (
    echo {"speakers":[],"headphones":[],"hotkeys":{"switch_device":"ctrl+alt+s","switch_type":"ctrl+alt+t"},"current_type":"Speakers", "kernel_mode_enabled": false, "force_start": false} > config.json
)

echo [%date% %time%] Verifying pystray installation... >> "%LOG_FILE%"
"%PYTHON_DIR%\python.exe" -c "import pystray; from PIL import Image; icon = pystray.Icon('test', Image.new('RGB', (32, 32), 'black')); print('Pystray OK')" >> "%LOG_FILE%" 2>&1
if %ERRORLEVEL% neq 0 (
    echo [%date% %time%] ERROR: Pystray test failed >> "%LOG_FILE%"
    echo Failed to verify pystray installation. Check the logs at: %LOG_FILE%
    pause
    exit /b 1
)

echo [%date% %time%] Launching application... >> "%LOG_FILE%"
echo [%date% %time%] Using Python: "%PYTHON_DIR%\pythonw.exe" >> "%LOG_FILE%"
echo [%date% %time%] Current directory: %CD% >> "%LOG_FILE%"
echo [%date% %time%] Files in directory: >> "%LOG_FILE%"
dir >> "%LOG_FILE%"

start /B "" "%PYTHON_DIR%\python.exe" audio_switcher.py >> "%LOG_FILE%" 2>&1

echo.
echo Audio Switcher is starting...
echo Detailed logs are being written to: %LOG_FILE%
echo Waiting for initialization...
timeout /t 5 > nul

echo [%date% %time%] Checking for process... >> "%LOG_FILE%"
tasklist | findstr "python" >> "%LOG_FILE%"

echo.
echo If no system tray icon appears, check the logs at:
echo %LOG_FILE%
echo.
echo Press any key to close this window (app will continue running)...
pause > nul
endlocal
