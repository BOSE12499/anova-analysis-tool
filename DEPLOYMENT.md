# ANOVA Analysis Tool - Deployment Guide

## 🚀 Quick Deployment to Render.com

### Version: v1.0.7
**Release Date**: September 4, 2025

### ✨ Latest Features
- Optimized chart sizes (9×5 inches) with proportional font scaling
- Colored p-values: 🟢 Green (non-significant ≥0.05) | 🔴 Red (significant <0.05)
- Enhanced variance analysis with optional pooled std dev line
- Production-ready configuration for Render.com

---

## 🛠️ Deployment Steps

### 1. GitHub Repository Setup
```bash
cd "c:\Users\freeb\Downloads\WEB Calculator"
git add .
git commit -m "Deploy: v1.0.7 - Production ready with optimized charts"
git tag -a v1.0.7 -m "Release v1.0.7: Optimized charts and colored p-values"
git push origin main
git push origin v1.0.7
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
- Flask 2.3.3
- pandas 2.1.1  
- numpy 1.24.3
- matplotlib 3.7.2
- scipy 1.11.3
- gunicorn 21.2.0

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

- **v1.0.7** (Current) - Production ready with optimized charts
- **v1.0.6** - PowerPoint export functionality
- **v1.0.5** - Enhanced deployment configuration

Ready for production! 🎉
