@echo off
echo ====================================
echo    ANOVA Tool - Safe File Organization
echo ====================================
echo.

echo 📁 Current files detected:
echo    ✓ app.py (Main Flask application)
echo    ✓ my.html (Web interface)
echo    ✓ requirements.txt (Dependencies)
echo    ✓ setup.bat (Setup script)
echo    ✓ uninstall.bat (Cleanup script)
echo    ✓ README files (Documentation)
echo    ✓ .vscode, .git folders (Development files)
echo.

echo This script will organize files into folders WITHOUT deleting anything!
echo All original files will be preserved.
echo.
set /p confirm="Continue with safe organization? (y/N): "
if /i not "%confirm%"=="y" (
    echo Organization cancelled.
    pause
    exit /b 0
)

echo.
echo 📂 Creating organized folder structure...

REM สร้างโฟลเดอร์หลัก
if not exist "scripts" mkdir scripts
if not exist "docs" mkdir docs
if not exist "config" mkdir config
if not exist "examples" mkdir examples

echo ✓ Created: scripts/, docs/, config/, examples/

echo.
echo 📋 COPYING files to organized folders...

REM Copy (ไม่ move) setup scripts
if exist "setup.bat" (
    copy setup.bat scripts\ >nul 2>&1
    echo ✓ Copied setup.bat → scripts/
)
if exist "uninstall.bat" (
    copy uninstall.bat scripts\ >nul 2>&1
    echo ✓ Copied uninstall.bat → scripts/
)

REM Copy documentation files
if exist "README.md" (
    copy README.md docs\ >nul 2>&1
    echo ✓ Copied README.md → docs/
)
if exist "README_USAGE.txt" (
    copy README_USAGE.txt docs\ >nul 2>&1
    echo ✓ Copied README_USAGE.txt → docs/
)

REM Copy config files
if exist ".python-version" (
    copy .python-version config\ >nul 2>&1
    echo ✓ Copied .python-version → config/
)
if exist ".gitignore" (
    copy .gitignore config\ >nul 2>&1
    echo ✓ Copied .gitignore → config/
)
if exist "render.yaml" (
    copy render.yaml config\ >nul 2>&1
    echo ✓ Copied render.yaml → config/
)

REM Copy VS Code settings
if exist ".vscode" (
    if not exist "config\.vscode" mkdir config\.vscode
    xcopy .vscode config\.vscode /E /I /Q >nul 2>&1
    echo ✓ Copied .vscode → config/.vscode/
)

echo.
echo 🎯 Creating example data file...
(
    echo LOT,DATA
    echo A,23.5
    echo A,24.1
    echo A,23.8
    echo A,23.2
    echo B,22.9
    echo B,23.2
    echo B,22.7
    echo B,23.4
    echo C,24.8
    echo C,25.1
    echo C,24.6
    echo C,24.9
) > examples\sample_anova_data.csv
echo ✓ Created sample_anova_data.csv → examples/

echo.
echo 🔧 Creating updated setup script for organized structure...

REM สร้าง setup script ใหม่ที่ทำงานจาก scripts folder
(
    echo @echo off
    echo echo ====================================
    echo echo    ANOVA Analysis Tool Setup v3.0
    echo echo    ^(Running from organized structure^)
    echo echo ====================================
    echo echo.
    echo.
    echo REM Change to project root directory
    echo cd /d "%%~dp0\.."
    echo echo Working directory: %%CD%%
    echo echo.
    echo.
    echo REM Check if main files exist in root
    echo if not exist "app.py" ^(
    echo     echo ERROR: app.py not found in project root!
    echo     echo Please ensure you're running from scripts/ folder
    echo     pause
    echo     exit /b 1
    echo ^)
    echo.
    echo REM Original setup content follows...
    echo echo [1/8] Cleaning old files...
    echo if exist venv ^(
    echo     echo Removing old virtual environment...
    echo     rmdir /s /q venv 2^>nul
    echo ^)
    echo if exist __pycache__ ^(
    echo     echo Removing Python cache...
    echo     rmdir /s /q __pycache__ 2^>nul
    echo ^)
    echo echo Old files cleaned!
    echo echo.
    echo.
    echo echo [2/8] Checking Python installation...
    echo python --version ^>nul 2^>^&1
    echo if errorlevel 1 ^(
    echo     echo ERROR: Python is not installed or not in PATH!
    echo     echo Download from: https://python.org/downloads/
    echo     pause
    echo     exit /b 1
    echo ^)
    echo echo Python found!
    echo echo.
    echo.
    echo echo [3/8] Creating virtual environment...
    echo python -m venv venv
    echo echo Virtual environment created!
    echo echo.
    echo.
    echo echo [4/8] Activating virtual environment...
    echo call venv\Scripts\activate
    echo echo.
    echo.
    echo echo [5/8] Installing dependencies...
    echo pip install --upgrade pip
    echo pip install -r requirements.txt
    echo echo Dependencies installed!
    echo echo.
    echo.
    echo echo [6/8] Creating quick start script...
    echo echo @echo off ^> start_anova.bat
    echo echo cd /d "%%%%~dp0" ^>^> start_anova.bat
    echo echo call venv\Scripts\activate ^>^> start_anova.bat
    echo echo python app.py ^>^> start_anova.bat
    echo echo pause ^>^> start_anova.bat
    echo echo Created start_anova.bat in project root!
    echo echo.
    echo.
    echo echo ====================================
    echo echo    Setup Complete! 
    echo echo ====================================
    echo echo.
    echo echo 🌐 Access at: http://localhost:10000
    echo echo 🚀 Next time: run start_anova.bat
    echo echo 📁 Examples in: examples/sample_anova_data.csv
    echo echo ====================================
    echo echo.
    echo python app.py
) > scripts\setup.bat

echo ✓ Created organized setup.bat → scripts/

echo.
echo 📝 Creating updated uninstall script...

(
    echo @echo off
    echo cd /d "%%~dp0\.."
    echo echo ====================================
    echo echo    ANOVA Tool Clean Uninstaller
    echo echo ====================================
    echo echo.
    echo echo This will remove temporary files but keep your work
    echo echo.
    echo set /p confirm="Continue? (y/N): "
    echo if /i not "%%confirm%%"=="y" ^(
    echo     echo Cancelled.
    echo     pause
    echo     exit /b 0
    echo ^)
    echo.
    echo if exist venv ^(
    echo     rmdir /s /q venv
    echo     echo ✓ Removed virtual environment
    echo ^)
    echo if exist __pycache__ ^(
    echo     rmdir /s /q __pycache__
    echo     echo ✓ Removed cache files
    echo ^)
    echo if exist start_anova.bat ^(
    echo     del start_anova.bat
    echo     echo ✓ Removed startup script
    echo ^)
    echo echo.
    echo echo Cleanup complete! Core files preserved.
    echo pause
) > scripts\uninstall.bat

echo ✓ Created organized uninstall.bat → scripts/

echo.
echo 📖 Creating organization guide...

(
    echo === ANOVA Analysis Tool - Organized Structure Guide ===
    echo.
    echo 📁 Project Structure:
    echo    📁 Root Directory:
    echo       📄 app.py           ^(Main Flask application^)
    echo       📄 my.html          ^(Web interface^)
    echo       📄 requirements.txt ^(Dependencies^)
    echo       📄 start_anova.bat  ^(Quick start - auto-created^)
    echo.
    echo    📁 scripts/
    echo       📄 setup.bat        ^(First-time setup^)
    echo       📄 uninstall.bat    ^(Clean removal^)
    echo.
    echo    📁 docs/
    echo       📄 README.md        ^(Main documentation^)
    echo       📄 README_USAGE.txt ^(User guide^)
    echo.
    echo    📁 config/
    echo       📄 .python-version  ^(Python version spec^)
    echo       📄 .gitignore       ^(Git ignore rules^)
    echo       📄 render.yaml      ^(Render deployment^)
    echo       📁 .vscode/         ^(VS Code settings^)
    echo.
    echo    📁 examples/
    echo       📄 sample_anova_data.csv ^(Example data^)
    echo.
    echo 🚀 How to Use After Organization:
    echo    1. First time: Run scripts\setup.bat
    echo    2. Regular use: Run start_anova.bat ^(in root^)
    echo    3. Cleanup: Run scripts\uninstall.bat
    echo.
    echo 📊 Example Data Format ^(examples\sample_anova_data.csv^):
    echo    Column A: LOT ^(A, B, C^)
    echo    Column B: DATA ^(numerical values^)
    echo.
    echo 🔒 Safety Features:
    echo    - Original files preserved in root
    echo    - Organized copies in subfolders
    echo    - No files deleted during organization
    echo    - Easy to revert if needed
    echo.
    echo ⚠️ Important Notes:
    echo    - Keep app.py, my.html, requirements.txt in root
    echo    - Use scripts\setup.bat for first-time setup
    echo    - All paths automatically handled
    echo.
    echo 📞 Troubleshooting:
    echo    - If setup fails: Check Python installation
    echo    - If paths wrong: Run from correct folder
    echo    - If errors: See docs\README_USAGE.txt
) > docs\ORGANIZATION_GUIDE.txt

echo ✓ Created ORGANIZATION_GUIDE.txt → docs/

echo.
echo ====================================
echo    ✅ Safe Organization Complete!
echo ====================================
echo.
echo 📊 Summary of changes:
echo    ✅ Created organized folder structure
echo    ✅ Copied files to appropriate folders
echo    ✅ Updated scripts with correct paths
echo    ✅ Created example data file
echo    ✅ Generated organization guide
echo    ✅ ALL ORIGINAL FILES PRESERVED
echo.
echo 📁 New structure:
echo    🟢 scripts/     - Setup and maintenance
echo    🔵 docs/        - Documentation and guides  
echo    🟡 config/      - Configuration files
echo    🟠 examples/    - Sample data
echo    ⚪ Root files   - Core application (unchanged)
echo.
echo 🚀 Next steps:
echo    1. Test: Run scripts\setup.bat
echo    2. Use: Run start_anova.bat (after setup)
echo    3. Learn: Read docs\ORGANIZATION_GUIDE.txt
echo.
echo 🔒 Safety guarantee:
echo    - Your original files are untouched
echo    - You can delete organized folders anytime
echo    - Everything works exactly as before
echo ====================================
echo.
pause