# 🚀 ANOVA Calculator - Complete Deployment Guide

## 📋 Overview
This guide will help you deploy the ANOVA Calculator web application to production using Render.com with automated version management.

## 🎯 Features
- **Statistical Analysis**: Complete One-way ANOVA with Tukey HSD post-hoc tests
- **Interactive UI**: Responsive web interface with dark/light theme
- **Production Ready**: Optimized Flask backend with health monitoring
- **Version Management**: Automated versioning with Git integration
- **Auto-deployment**: Render.com integration with GitHub

## 🛠️ Prerequisites
1. **Python 3.9+** installed
2. **Git** installed and configured
3. **GitHub account**
4. **Render.com account** (free tier available)

## 📦 Quick Start

### Step 1: Prepare Your Repository
```bash
# Windows users: run deploy.bat
# Linux/Mac users: run deploy.sh

# Or manually:
git init
git add .
git commit -m "Initial commit: ANOVA Calculator v1.0.0"
```

### Step 2: Version Management
Use the automated version management system:

```bash
# For bug fixes (1.0.0 → 1.0.1)
python update_version.py patch "Fix calculation precision"

# For new features (1.0.0 → 1.1.0) 
python update_version.py minor "Add export functionality"

# For breaking changes (1.0.0 → 2.0.0)
python update_version.py major "New UI framework"
```

### Step 3: Deploy to Render.com

#### Option A: GitHub Integration (Recommended)
1. **Push to GitHub:**
   ```bash
   git remote add origin https://github.com/yourusername/anova-calculator.git
   git push -u origin main
   ```

2. **Connect to Render.com:**
   - Go to [render.com](https://render.com)
   - Click "New +" → "Web Service"
   - Connect your GitHub repository
   - Render will auto-detect the `render.yaml` configuration

#### Option B: Direct Deploy
1. **Manual Setup on Render:**
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `gunicorn --bind 0.0.0.0:$PORT app:app`
   - Environment: Python 3.9+

## 🔧 Configuration Files

### render.yaml
```yaml
services:
  - type: web
    name: anova-calculator
    env: python
    buildCommand: pip install -r requirements.txt
    startCommand: gunicorn --bind 0.0.0.0:$PORT --workers 2 --timeout 120 app:app
    healthCheckPath: /health
    envVars:
      - key: PYTHONUNBUFFERED
        value: "1"
      - key: VERSION
        sync: false
```

### requirements.txt
```txt
Flask==2.3.3
flask-cors==4.0.0
pandas==2.1.4
numpy==1.24.4
scipy==1.11.4
matplotlib==3.8.2
pingouin==0.5.4
statsmodels==0.14.1
gunicorn==21.2.0
```

## 📊 Version History Tracking

The system automatically maintains version history in `VERSION.txt`:

```
Version: v1.0.0
Release Date: 2024-01-15
Release Type: Initial Release

Features:
- Complete One-way ANOVA analysis
- Tukey HSD post-hoc tests
- Interactive data visualization
- Responsive web interface
- CSV file support
- Statistical summary tables

Dependencies:
- Python 3.9+
- Flask 2.3.3
- pandas, numpy, scipy
- matplotlib for plotting
```

## 🔍 Health Monitoring

The application includes built-in health checks:
- **Endpoint**: `/health`
- **Version Info**: `/version`
- **Status**: Returns server status and version information

## 🚀 Auto-deployment Workflow

1. **Make Changes** to your code
2. **Version Update**:
   ```bash
   python update_version.py patch "Your change description"
   ```
3. **Push to GitHub**:
   ```bash
   git push origin main
   ```
4. **Automatic Deploy**: Render.com detects changes and redeploys

## 🐛 Troubleshooting

### Common Issues:

1. **Build Failures**:
   - Check `requirements.txt` versions
   - Verify Python version compatibility
   - Check render.yaml syntax

2. **Memory Issues**:
   - Optimized for Render.com free tier (512MB RAM)
   - Plot generation uses memory optimization
   - Garbage collection after each analysis

3. **Version Conflicts**:
   - Use `git status` to check repository state
   - Ensure `VERSION.txt` is committed
   - Verify version format (vX.Y.Z)

### Debug Commands:
```bash
# Check current version
cat VERSION.txt

# Test application locally
python app.py

# Check dependencies
pip list

# Validate render.yaml
render-cli validate render.yaml
```

## 📈 Performance Optimization

### Free Tier Optimizations:
- **Memory**: Aggressive plot cleanup and garbage collection
- **CPU**: Optimized statistical calculations
- **Network**: Compressed responses and efficient routing
- **Storage**: Minimal file I/O operations

### Production Optimizations:
- Gunicorn with 2 workers
- Request timeout: 120 seconds
- Health check monitoring
- Environment variable configuration

## 🔐 Security Features

- CORS enabled for cross-origin requests
- Content-Type validation
- File size limits (16MB max)
- Input sanitization and validation
- Error handling without sensitive data exposure

## 📝 API Endpoints

| Endpoint | Method | Description |
|----------|---------|------------|
| `/` | GET | Main application interface |
| `/analyze_anova` | POST | Statistical analysis endpoint |
| `/health` | GET | Health check status |
| `/version` | GET | Version information |
| `/dashboard` | GET | Analysis dashboard |

## 🎯 Usage Examples

### Statistical Analysis:
1. Upload CSV file with LOT and DATA columns
2. Set optional LSL/USL limits  
3. Click "Analyze" for complete ANOVA results
4. Export results as JSON/PDF

### Version Management:
```bash
# Patch release (bug fix)
python update_version.py patch "Fixed precision in F-statistic calculation"

# Minor release (new feature)  
python update_version.py minor "Added Bartlett test for variance equality"

# Major release (breaking change)
python update_version.py major "Updated to Flask 3.0 with new API structure"
```

## 🎉 Success Indicators

✅ **Deployment Successful** when you see:
- Health check returns `{"status": "OK"}`
- Version endpoint shows current version
- Main interface loads without errors
- Statistical analysis processes correctly

## 📞 Support

For issues or questions:
1. Check the troubleshooting section
2. Verify configuration files
3. Test locally before deploying
4. Check Render.com deployment logs

---

**Happy Analyzing! 📊✨**

*ANOVA Calculator v1.0.0 - Production Ready*
