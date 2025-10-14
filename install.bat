@echo off
echo ========================================
echo    Auto Tay Luyen Tool - Installation
echo ========================================
echo.

echo Checking Python installation...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.7+ from https://python.org
    pause
    exit /b 1
)

echo Python found. Installing required packages...
echo.

pip install pillow
pip install pyautogui
pip install pytesseract
pip install pygetwindow
pip install keyboard

echo.
echo ========================================
echo Installation completed!
echo.
echo IMPORTANT: You also need to install Tesseract-OCR:
echo 1. Download from: https://github.com/UB-Mannheim/tesseract/wiki
echo 2. Install with "Add to PATH" option checked
echo.
echo Then run: python auto_tay_luyen.py
echo ========================================
pause
