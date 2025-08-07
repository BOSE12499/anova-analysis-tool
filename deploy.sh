#!/bin/bash

# ANOVA Calculator Deployment Guide
# This script helps you deploy the ANOVA Calculator to Render.com

echo "🚀 ANOVA Calculator Deployment Setup"
echo "======================================"

# Check if git is initialized
if [ ! -d ".git" ]; then
    echo "📁 Initializing Git repository..."
    git init
    git branch -M main
fi

# Check if files are staged
if [ -z "$(git diff --cached --name-only)" ]; then
    echo "📋 Staging files for commit..."
    git add .
fi

# Check for version updates
if [ -f "update_version.py" ]; then
    echo "🏷️  Available version commands:"
    echo "   python update_version.py patch 'Bug fixes'"
    echo "   python update_version.py minor 'New features'"
    echo "   python update_version.py major 'Breaking changes'"
    echo ""
    echo "Run version script? (y/n)"
    read -r response
    if [[ "$response" =~ ^([yY][eE][sS]|[yY])$ ]]; then
        echo "Enter version type (patch/minor/major):"
        read -r version_type
        echo "Enter commit message:"
        read -r commit_msg
        python update_version.py "$version_type" "$commit_msg"
    fi
fi

# Check if we need to commit
if [ -n "$(git diff --cached --name-only)" ] || [ -n "$(git status --porcelain)" ]; then
    echo "💾 Creating commit..."
    git commit -m "Deploy: Ready for production deployment"
fi

echo ""
echo "🌐 Next steps for Render.com deployment:"
echo "1. Push to GitHub: git remote add origin <your-repo-url>"
echo "2. Push code: git push -u origin main"
echo "3. Connect your GitHub repo to Render.com"
echo "4. Use render.yaml configuration for auto-deployment"
echo ""
echo "📝 render.yaml is already configured with:"
echo "   - Health check endpoint: /health"
echo "   - Environment variables"
echo "   - Production Gunicorn settings"
echo ""
echo "✅ Deployment setup complete!"
