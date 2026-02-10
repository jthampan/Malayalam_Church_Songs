@echo off
REM Build Malayalam Church Songs PPT Generator
REM Windows Batch Script

REM Change to the directory where this batch file is located
cd /d "%~dp0"

echo.
echo ========================================
echo Building Church Songs PPT Generator
echo ========================================
echo.
echo Working directory: %CD%
echo.

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed!
    echo Please install Python from https://www.python.org/downloads/
    pause
    exit /b 1
)

REM Check if build_exe.py exists
if not exist "build_exe.py" (
    echo ERROR: build_exe.py not found in current directory!
    echo Please make sure you're running this from the Church_Songs folder.
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
