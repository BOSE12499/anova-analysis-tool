@echo off
REM ANOVA Calculator Deployment Setup for Windows

echo 🚀 ANOVA Calculator Deployment Setup
echo ======================================

REM Check if git is initialized
if not exist ".git" (
    echo 📁 Initializing Git repository...
    git init
    git branch -M main
)

REM Check if files are staged
for /f %%i in ('git diff --cached --name-only') do set staged=%%i
if "%staged%"=="" (
    echo 📋 Staging files for commit...
    git add .
)

REM Check for version updates
if exist "update_version.py" (
    echo 🏷️  Available version commands:
    echo    python update_version.py patch "Bug fixes"
    echo    python update_version.py minor "New features" 
    echo    python update_version.py major "Breaking changes"
    echo.
    set /p response="Run version script? (y/n): "
    if /i "%response%"=="y" (
        set /p version_type="Enter version type (patch/minor/major): "
        set /p commit_msg="Enter commit message: "
        python update_version.py "%version_type%" "%commit_msg%"
    )
)

REM Check if we need to commit
git status --porcelain > temp_status.txt
for /f %%i in (temp_status.txt) do set need_commit=true
del temp_status.txt

if defined need_commit (
    echo 💾 Creating commit...
    git commit -m "Deploy: Ready for production deployment"
)

echo.
echo 🌐 Next steps for Render.com deployment:
echo 1. Push to GitHub: git remote add origin ^<your-repo-url^>
echo 2. Push code: git push -u origin main
echo 3. Connect your GitHub repo to Render.com  
echo 4. Use render.yaml configuration for auto-deployment
echo.
echo 📝 render.yaml is already configured with:
echo    - Health check endpoint: /health
echo    - Environment variables
echo    - Production Gunicorn settings
echo.
echo ✅ Deployment setup complete!
pause
