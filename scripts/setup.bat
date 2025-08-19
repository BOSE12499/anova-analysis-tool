@echo off
echo ====================================
echo    ANOVA Analysis Tool Setup v2.0
echo ====================================
echo.

echo [1/8] Cleaning old files...
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

echo [2/8] Checking Python installation...
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH!
    echo Please install Python 3.11+ and add it to PATH
    echo Download from: https://python.org/downloads/
    pause
    exit /b 1
)
for /f "tokens=2" %%i in ('python --version') do set PYTHON_VERSION=%%i
echo Python %PYTHON_VERSION% found!
echo.

echo [3/8] Checking required files...
if not exist "app.py" (
    echo ERROR: app.py not found!
    echo Please make sure you're in the correct directory
    pause
    exit /b 1
)
if not exist "requirements.txt" (
    echo ERROR: requirements.txt not found!
    echo Please make sure you have the requirements file
    pause
    exit /b 1
)
if not exist "my.html" (
    echo ERROR: my.html not found!
    echo Please make sure you have the HTML file
    pause
    exit /b 1
)
echo Required files found: app.py, requirements.txt, my.html
echo.

echo [4/8] Checking internet connection...
ping google.com -n 1 >nul 2>&1
if errorlevel 1 (
    echo WARNING: No internet connection detected!
    echo Package installation may fail
    echo Please check your internet connection
    pause
)
echo Internet connection OK!
echo.

echo [5/8] Creating virtual environment...
python -m venv venv
if errorlevel 1 (
    echo ERROR: Failed to create virtual environment!
    pause
    exit /b 1
)
echo Virtual environment created!
echo.

echo [6/8] Activating virtual environment...
call venv\Scripts\activate
if errorlevel 1 (
    echo ERROR: Failed to activate virtual environment!
    pause
    exit /b 1
)
echo Virtual environment activated!
echo.

echo [7/8] Installing dependencies...
echo This may take a few minutes for scientific packages...
pip install --upgrade pip
pip install -r requirements.txt
if errorlevel 1 (
    echo ERROR: Failed to install dependencies!
    echo Please check your requirements.txt file and internet connection
    pause
    exit /b 1
)
echo Dependencies installed successfully!
echo - Flask (Web framework)
echo - scipy, pandas, numpy (Statistical analysis)
echo - matplotlib (Plotting)
echo - statsmodels (Advanced statistics)
echo.

echo [7.5/8] Creating startup shortcut...
echo @echo off > start_anova.bat
echo echo Starting ANOVA Analysis Tool... >> start_anova.bat
echo call venv\Scripts\activate >> start_anova.bat  
echo python app.py >> start_anova.bat
echo pause >> start_anova.bat
echo Created start_anova.bat for future quick startup!
echo.

echo [8/8] Checking port availability...
netstat -ano | findstr :10000 >nul
if not errorlevel 1 (
    echo WARNING: Port 10000 is already in use!
    echo Please close any applications using port 10000
    echo or the application may fail to start
    pause
)
echo Port 10000 is available!
echo.

echo ====================================
echo    Setup Complete Successfully! 
echo ====================================
echo.
echo 🎯 Your ANOVA Analysis Tool is ready!
echo.
echo 📊 Features included:
echo    - One-way ANOVA analysis
echo    - Tukey-Kramer HSD post-hoc tests
echo    - Variance equality tests (Levene, Bartlett, etc.)
echo    - Statistical plots and charts
echo    - CSV file upload and analysis
echo.
echo 🌐 Usage:
echo    1. Access at: http://localhost:10000
echo    2. Upload your CSV file (LOT in column A, DATA in column B)
echo    3. Optionally set LSL/USL limits
echo    4. Click "Perform Analysis"
echo.
echo 🚀 Next time usage:
echo    Just double-click: start_anova.bat
echo    (No need to run setup.bat again)
echo.
echo 📝 File structure expected:
echo    Column A: LOT (e.g., A, B, C)
echo    Column B: DATA (numerical values)
echo.
echo Press Ctrl+C to stop the server when done
echo ====================================
echo.

python app.py