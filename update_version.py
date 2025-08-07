#!/usr/bin/env python3
"""
Version Management Script for ANOVA Analysis Tool
Usage: python update_version.py [major|minor|patch] [optional_message]
"""

import re
import sys
import subprocess
from datetime import datetime
import os

def get_current_version():
    """Read current version from VERSION.txt"""
    try:
        with open('VERSION.txt', 'r', encoding='utf-8') as f:
            content = f.read()
            match = re.search(r'v(\d+\.\d+\.\d+)', content)
            if match:
                return match.group(1)
    except FileNotFoundError:
        return "0.0.0"
    return "0.0.0"

def increment_version(current_version, bump_type):
    """Increment version based on type"""
    major, minor, patch = map(int, current_version.split('.'))
    
    if bump_type == 'major':
        major += 1
        minor = 0
        patch = 0
    elif bump_type == 'minor':
        minor += 1
        patch = 0
    elif bump_type == 'patch':
        patch += 1
    
    return f"{major}.{minor}.{patch}"

def update_version_file(new_version, changes="Bug fixes and improvements"):
    """Update VERSION.txt with new version"""
    current_date = datetime.now().strftime("%Y-%m-%d")
    
    version_entry = f"""
v{new_version} ({current_date})
{'-' * (len(new_version) + 12)}
✅ {changes}
"""
    
    try:
        with open('VERSION.txt', 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Insert new version after the header
        lines = content.split('\n')
        
        # Find the end of the header section
        header_end_idx = 0
        for i, line in enumerate(lines):
            if line.startswith('v') and ('(' in line and ')' in line):
                header_end_idx = i
                break
        
        # Insert new version entry
        new_lines = lines[:header_end_idx] + version_entry.strip().split('\n') + [''] + lines[header_end_idx:]
        new_content = '\n'.join(new_lines)
        
        with open('VERSION.txt', 'w', encoding='utf-8') as f:
            f.write(new_content)
            
        print(f"✅ Updated VERSION.txt to v{new_version}")
        
    except Exception as e:
        print(f"❌ Error updating VERSION.txt: {e}")

def update_render_yaml_version(new_version):
    """Update version in render.yaml"""
    try:
        with open('render.yaml', 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Update VERSION environment variable
        updated_content = re.sub(
            r'(\s+- key: VERSION\s+value: ")[^"]+(")',
            f'\\1{new_version}\\2',
            content
        )
        
        with open('render.yaml', 'w', encoding='utf-8') as f:
            f.write(updated_content)
        
        print(f"✅ Updated render.yaml version to {new_version}")
        
    except Exception as e:
        print(f"❌ Error updating render.yaml: {e}")

def git_commit_and_tag(version, message):
    """Git commit and create tag"""
    try:
        # Add all changes
        subprocess.run(['git', 'add', '.'], check=True, capture_output=True)
        
        # Commit
        commit_msg = f"v{version} - {message}"
        subprocess.run(['git', 'commit', '-m', commit_msg], check=True, capture_output=True)
        
        # Create tag
        tag_msg = f"Version {version}: {message}"
        subprocess.run(['git', 'tag', '-a', f"v{version}", '-m', tag_msg], check=True, capture_output=True)
        
        # Push to origin
        subprocess.run(['git', 'push', 'origin', 'main'], check=True, capture_output=True)
        subprocess.run(['git', 'push', 'origin', '--tags'], check=True, capture_output=True)
        
        print(f"✅ Committed and tagged v{version}")
        print(f"📦 Pushed to GitHub with tag v{version}")
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"❌ Git error: {e}")
        if e.output:
            print(f"Error output: {e.output.decode()}")
        return False

def check_git_status():
    """Check if there are uncommitted changes"""
    try:
        result = subprocess.run(['git', 'status', '--porcelain'], 
                              capture_output=True, text=True, check=True)
        return len(result.stdout.strip()) == 0
    except:
        return False

def main():
    print("🚀 ANOVA Analysis Tool - Version Management")
    print("=" * 50)
    
    if len(sys.argv) < 2:
        print("Usage: python update_version.py [major|minor|patch] [optional_message]")
        print("\nExamples:")
        print("  python update_version.py patch \"Fixed responsive notifications\"")
        print("  python update_version.py minor \"Added new statistical tests\"")
        print("  python update_version.py major \"Complete UI redesign\"")
        sys.exit(1)
    
    bump_type = sys.argv[1].lower()
    if bump_type not in ['major', 'minor', 'patch']:
        print("❌ Error: bump_type must be 'major', 'minor', or 'patch'")
        sys.exit(1)
    
    message = sys.argv[2] if len(sys.argv) > 2 else "Updates and improvements"
    
    # Check if we're in a git repository
    if not os.path.exists('.git'):
        print("❌ Error: Not in a Git repository")
        sys.exit(1)
    
    current_version = get_current_version()
    new_version = increment_version(current_version, bump_type)
    
    print(f"📊 Current version: v{current_version}")
    print(f"🔄 New version: v{new_version}")
    print(f"📝 Message: {message}")
    print()
    
    # Confirm with user
    confirm = input("Continue with version update? (y/N): ").lower().strip()
    if confirm not in ['y', 'yes']:
        print("❌ Operation cancelled")
        sys.exit(0)
    
    # Update version files
    print("📝 Updating version files...")
    update_version_file(new_version, message)
    update_render_yaml_version(new_version)
    
    # Git operations
    print("📦 Creating Git commit and tag...")
    if git_commit_and_tag(new_version, message):
        print()
        print("🎉 Successfully released v{new_version}!")
        print(f"📦 GitHub Release: https://github.com/BOSE12499/anova-analysis-tool/releases/tag/v{new_version}")
        print(f"🌐 Deploy will start automatically on Render")
        print(f"🔗 Live URL: https://anova-analysis-tool.onrender.com")
        print()
        print("⏳ Deployment typically takes 2-3 minutes...")
        print("💡 Check deployment status at: https://dashboard.render.com")
    else:
        print("❌ Failed to create release")
        sys.exit(1)

if __name__ == "__main__":
    main()
