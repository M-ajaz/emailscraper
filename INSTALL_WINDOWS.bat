@echo off
echo ============================================
echo  Mail Scraper — Windows Setup Script
echo ============================================
echo.

REM Check if Python 3.11 is installed
python --version 2>nul | findstr "3.11" >nul
if errorlevel 1 (
    echo ERROR: Python 3.11 is not installed or not in PATH.
    echo.
    echo Please:
    echo 1. Download Python 3.11 from:
    echo    https://www.python.org/downloads/release/python-3119/
    echo 2. During install, CHECK "Add Python to PATH"
    echo 3. Re-run this script after installing
    echo.
    pause
    exit /b 1
)

echo Python 3.11 found. Continuing...
echo.

REM Create .env if not exists
if not exist "backend\.env" (
    echo Creating .env from template...
    copy "backend\.env.template" "backend\.env"
    echo.
    echo IMPORTANT: Edit backend\.env with your email credentials before launching!
    echo Open it now? [Y/N]
    set /p open=
    if /i "%open%"=="Y" notepad "backend\.env"
)

REM Install Python packages
echo Installing Python packages...
cd backend
pip install -r requirements.txt --no-cache-dir
if errorlevel 1 (
    echo.
    echo Package install failed. Try running as Administrator.
    pause
    exit /b 1
)
cd ..

REM Install Node packages
echo Installing Node.js packages...
cd frontend
npm install
cd ..

cd electron
npm install
cd ..

echo.
echo ============================================
echo  Setup Complete!
echo ============================================
echo.
echo To run in development mode:
echo   cd backend ^&^& python -m uvicorn main:app --port 8000
echo   cd frontend ^&^& npm run dev
echo   cd electron ^&^& npm run dev
echo.
echo To build installer:
echo   cd electron ^&^& npm run build:win
echo.
pause
