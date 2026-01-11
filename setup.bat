@echo off
setlocal
echo Checking for Python...

:: Try 'py' launcher first (standard on Windows)
py --version >nul 2>&1
if %errorlevel% == 0 (
    set PY_CMD=py
    goto :found
)

:: Try 'python' command
python --version >nul 2>&1
if %errorlevel% == 0 (
    set PY_CMD=python
    goto :found
)

echo [!] ERROR: Python was not found. 
echo Please download and install Python from: https://www.python.org/downloads/
echo *** IMPORTANT: Check the "Add Python to PATH" box during installation! ***
pause
exit /b

:found
echo Using %PY_CMD% to set up environment...
%PY_CMD% -m venv venv
if %errorlevel% neq 0 (
    echo [!] Failed to create virtual environment.
    pause
    exit /b
)

call venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt
echo.
echo Setup complete. Run "start_app.bat" to launch.
pause
