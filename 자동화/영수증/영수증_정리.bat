@echo off
chcp 65001 > nul

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python is not installed.
    echo Please install Python 3.x from: https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation.
    pause
    exit /b 1
)

echo Installing required packages...
pip install pypdf pdfplumber reportlab openpyxl pymupdf pillow --quiet --disable-pip-version-check

echo Installing Windows OCR wrapper (lightweight, no model download)...
pip install winocr --quiet --disable-pip-version-check

echo.
echo Starting...
echo.
python "%~dp0receipt_organizer.py"

if %errorlevel% neq 0 (
    pause
    exit /b 1
)
