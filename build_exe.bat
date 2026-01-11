@echo off
echo ========================================
echo Building Shift Automator Pro (Portable)
echo ========================================
if not exist venv (
    echo [ERROR] Virtual environment not found. Run setup.bat first.
    pause
    exit /b
)
call venv\Scripts\activate
pip install -r requirements.txt
pyinstaller --noconsole --onefile --name "Shift Automator Pro" --icon=icon.ico main.py
echo.
echo [SUCCESS] Build complete! Check the "dist" folder.
pause
