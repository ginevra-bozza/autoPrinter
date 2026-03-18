@echo off
TITLE Automatic Printer
COLOR 0A

echo Checking if Python is installed...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python is not installed! 
    echo Please install Python from the Windows Store or python.org.
    pause
    exit /b
)

echo Installing required background tools...
pip install imap-tools pillow img2pdf pywin32 --quiet

echo Starting the printer listener...
echo DO NOT CLOSE THIS WINDOW. You can minimize it.
python autoPrinter.py
pause