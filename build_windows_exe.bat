@echo off
REM Build Malayalam Church Songs PPT Generator
REM Windows Batch Script

echo.
echo ========================================
echo Building Church Songs PPT Generator
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed!
    echo Please install Python from https://www.python.org/downloads/
    pause
    exit /b 1
)

echo Installing PyInstaller...
pip install pyinstaller python-pptx pandas openpyxl

echo.
echo Building executable...
python build_exe.py

echo.
echo ========================================
echo Build Complete!
echo ========================================
echo.
echo Your executable is in the 'dist' folder:
echo   dist\Church_Songs_Generator.exe
echo.
echo You can distribute this file to church members.
echo No Python installation needed for end users!
echo.
pause
