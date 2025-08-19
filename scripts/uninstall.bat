@echo off
echo ====================================
echo    ANOVA Analysis Tool Uninstaller
echo ====================================
echo.

echo This will remove the ANOVA Analysis Tool environment
echo but keep your original files (app.py, my.html, etc.)
echo.
set /p confirm="Are you sure you want to uninstall? (y/N): "
if /i not "%confirm%"=="y" (
    echo Uninstall cancelled.
    pause
    exit /b 0
)

echo.
echo Removing virtual environment...
if exist venv (
    rmdir /s /q venv
    echo ✓ Virtual environment removed!
) else (
    echo - No virtual environment found.
)

echo Removing cache files...
if exist __pycache__ (
    rmdir /s /q __pycache__
    echo ✓ Python cache files removed!
) else (
    echo - No cache files found.
)

echo Removing shortcuts...
if exist start_anova.bat (
    del start_anova.bat
    echo ✓ Startup shortcut removed!
) else (
    echo - No startup shortcut found.
)

echo Removing temporary files...
if exist runtime.txt (
    del runtime.txt
    echo ✓ Runtime file removed!
)

echo.
echo ====================================
echo    Uninstall Complete!
echo ====================================
echo.
echo The following files have been preserved:
echo ✓ app.py (Main application)
echo ✓ my.html (Web interface)
echo ✓ requirements.txt (Dependencies list)
echo ✓ setup.bat (Setup script)
echo ✓ README.md (Documentation)
echo ✓ Any CSV files you created
echo.
echo To reinstall: Run setup.bat again
echo ====================================
pause