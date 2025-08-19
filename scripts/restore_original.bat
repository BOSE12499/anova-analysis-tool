@echo off
echo ====================================
echo    Restore Original File Structure
echo ====================================
echo.

echo This will remove organized folders and restore original structure
echo.
echo Folders to be removed (if they exist):
echo - scripts/
echo - docs/
echo - config/
echo - examples/
echo.
echo Original files will remain untouched:
echo ✓ app.py
echo ✓ my.html 
echo ✓ requirements.txt
echo ✓ setup.bat
echo ✓ uninstall.bat
echo ✓ README files
echo.

set /p confirm="Continue with restoration? (y/N): "
if /i not "%confirm%"=="y" (
    echo Restoration cancelled.
    pause
    exit /b 0
)

echo.
echo Removing organized folders...

if exist "scripts" (
    rmdir /s /q scripts
    echo ✓ Removed scripts/
) else (
    echo - scripts/ not found
)

if exist "docs" (
    rmdir /s /q docs
    echo ✓ Removed docs/
) else (
    echo - docs/ not found
)

if exist "config" (
    rmdir /s /q config
    echo ✓ Removed config/
) else (
    echo - config/ not found
)

if exist "examples" (
    rmdir /s /q examples
    echo ✓ Removed examples/
) else (
    echo - examples/ not found
)

echo.
echo ====================================
echo    Restoration Complete!
echo ====================================
echo.
echo Your project is back to original structure:
echo 📄 app.py
echo 📄 my.html
echo 📄 requirements.txt
echo 📄 setup.bat
echo 📄 uninstall.bat
echo 📄 README files
echo 📁 .vscode/, .git/ (unchanged)
echo.
echo You can now use setup.bat normally
echo ====================================
pause