# ANOVA Analysis Tool - Deployment Guide

## 🚀 Quick Deployment to Render.com

### Version: v1.0.9
**Release Date**: September 4, 2025

### 🔧 **CRITICAL FIXES - v1.0.9**
- **✅ Resolved Python 3.13 compilation errors** with binary-only installation
- **✅ Force runtime: python-3.11.8** in render.yaml
- **✅ Binary wheels only** (--only-binary=all) to avoid source compilation
- **✅ Enhanced build configuration** with proper pip settings

---

## 🛠️ Deployment Steps

### 1. GitHub Repository Setup
```bash
cd "c:\Users\freeb\Downloads\WEB Calculator"
git add .
git commit -m "Deploy: v1.0.9 - CRITICAL FIX for Python 3.13 compilation errors"
git tag -a v1.0.9 -m "Release v1.0.9: Fixed pandas compilation with binary-only installation"
git push origin main
git push origin v1.0.9
```

### 2. Render.com Deployment
1. Go to [render.com](https://render.com)
2. Create new **Web Service**
3. Connect GitHub repository: `BOSE12499/anova-analysis-tool`
4. Use **Automatic deploys from Git**
5. Render will use `config/render.yaml` automatically

### 3. Configuration Files Ready
- ✅ **render.yaml** - Render configuration
- ✅ **requirements.txt** - Python dependencies  
- ✅ **Health check** - `/health` endpoint
- ✅ **GitHub Actions** - CI/CD pipeline

---

## 📋 Configuration Details

### Render.yaml Settings
```yaml
services:
  - type: web
    name: anova-analysis-tool
    env: python
    buildCommand: "pip install -r requirements.txt"
    startCommand: "gunicorn --bind 0.0.0.0:$PORT --workers 1 --timeout 60 src.app:app"
    healthCheckPath: "/health"
```

### Dependencies (requirements.txt)
- Flask 3.0.0
- pandas 2.2.0 (Python 3.11 compatible)
- numpy 1.26.3
- matplotlib 3.8.2
- scipy 1.12.0
- gunicorn 21.2.0

### Runtime Configuration
- **runtime.txt**: python-3.11.8 (forced version)
- **Build Command**: pip install --upgrade pip && pip install -r requirements.txt

---

## 🔗 Post-Deployment

After successful deployment, your ANOVA Analysis Tool will be available at:
`https://your-app-name.onrender.com`

### Available Endpoints
- `/` - Main application
- `/health` - Health check
- `/version` - Version information
- `/analyze_anova` - API endpoint

---

## 📊 Features Overview

### Statistical Analysis
- One-way ANOVA with F-test
- Tukey HSD post-hoc comparisons
- Variance equality tests (Levene, Brown-Forsythe, Bartlett)
- Welch ANOVA for unequal variances

### Visualizations
- Box plots with group means (🟢 green markers)
- Tukey HSD comparison charts
- Variance analysis scatter plots
- Professional styling with optimized sizing

### Export Options
- JSON data export
- PowerPoint presentations
- Statistical tables with colored p-values

---

## 🏷️ Version Tags

- **v1.0.8** (Current) - Fixed Python 3.13 compatibility for Render
- **v1.0.7** - Production ready with optimized charts
- **v1.0.6** - PowerPoint export functionality
- **v1.0.5** - Enhanced deployment configuration

Ready for production! 🎉
