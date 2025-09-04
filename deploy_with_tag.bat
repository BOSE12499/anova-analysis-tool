@echo off
echo 🚀 ANOVA Analysis Tool - Deployment Script
echo ==========================================

REM Update version and create new tag
echo 📝 Updating version...
python scripts\update_version.py minor "Deploy: Optimized charts, colored p-values, ready for Render deployment"

REM Add all files
echo 📋 Adding files to Git...
git add .

REM Check if there are changes to commit
git diff --cached --quiet
if errorlevel 1 (
    echo 💾 Creating commit...
    git commit -m "Deploy: v1.0.7 - Optimized charts and colored p-values for production"
    
    REM Get the version for tagging
    for /f "tokens=*" %%i in ('type docs\VERSION.txt ^| findstr /r "^v[0-9]"') do set VERSION=%%i
    set VERSION=%VERSION:~0,6%
    
    echo 🏷️  Creating Git tag: %VERSION%
    git tag -a %VERSION% -m "Release %VERSION%: Optimized charts, colored p-values, ready for Render deployment"
    
    echo 🌐 Pushing to GitHub...
    git push origin main
    git push origin %VERSION%
    
    echo ✅ Deployment complete!
    echo 📍 Tag %VERSION% created and pushed to GitHub
    echo 🔗 Ready for Render deployment
) else (
    echo ℹ️  No changes to commit
)

pause
