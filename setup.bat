@echo off
echo ====================================
echo    ANOVA Analysis Tool Setup
echo ====================================
echo.

echo [1/6] Cleaning old files...
if exist venv (
    echo Removing old virtual environment...
    rmdir /s /q venv 2>nul
)
if exist __pycache__ (
    echo Removing Python cache...
    rmdir /s /q __pycache__ 2>nul
)
if exist runtime.txt (
    echo Removing empty runtime.txt...
    del runtime.txt 2>nul
)
echo Old files cleaned!
echo.

echo [2/6] Checking Python installation...
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH!
    echo Please install Python 3.11+ and add it to PATH
    pause
    exit /b 1
)
python --version
echo.

echo [3/6] Creating virtual environment...
python -m venv venv
if errorlevel 1 (
    echo ERROR: Failed to create virtual environment!
    pause
    exit /b 1
)
echo Virtual environment created!
echo.

echo [4/6] Activating virtual environment...
call venv\Scripts\activate
if errorlevel 1 (
    echo ERROR: Failed to activate virtual environment!
    pause
    exit /b 1
)
echo Virtual environment activated!
echo.

echo [5/6] Installing dependencies...
echo This may take a few minutes...
pip install --upgrade pip
pip install -r requirements.txt
if errorlevel 1 (
    echo ERROR: Failed to install dependencies!
    echo Please check your requirements.txt file
    pause
    exit /b 1
)
echo Dependencies installed successfully!
echo.

echo [6/6] Testing application...
echo Starting Flask application...
echo.
echo ====================================
echo    Setup Complete! 
echo ====================================
echo.
echo Your ANOVA Analysis Tool is ready!
echo Access it at: http://localhost:10000
echo.
echo Press Ctrl+C to stop the server
echo ====================================
echo.

python app.py