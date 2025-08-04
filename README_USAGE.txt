=== ANOVA Analysis Tool - Quick Start Guide ===

Prerequisites:
- Windows 10/11
- Python 3.11 or higher (from python.org)
- Internet connection for downloading packages

First Time Setup:
1. Double-click setup.bat
2. Wait for setup to complete (may take 2-5 minutes)
3. Browser will open automatically at http://localhost:10000

How to Use:
1. Prepare your CSV file with:
   - Column A: LOT (categories like A, B, C)
   - Column B: DATA (numerical values)
2. Upload CSV file on the webpage
3. Optionally set LSL/USL specification limits
4. Click "Perform Analysis"
5. View comprehensive ANOVA results

Features Included:
✓ One-way ANOVA analysis
✓ Tukey-Kramer HSD post-hoc tests
✓ Variance equality tests (Levene, Bartlett, Brown-Forsythe, O'Brien)
✓ Statistical plots and charts
✓ Mean absolute deviation analysis
✓ Connecting letters report
✓ Professional statistical output (JMP-style)

Next Time Usage:
- Just double-click: start_anova.bat
- No need to run setup.bat again

Troubleshooting:
- Python not found: Install from https://python.org/downloads/
- Port 10000 busy: Close other applications or restart computer
- Package install fails: Check internet connection
- Analysis errors: Ensure CSV has correct format (LOT, DATA columns)

File Structure:
- app.py: Main Flask application
- my.html: Web interface
- requirements.txt: Python dependencies
- setup.bat: First-time setup
- start_anova.bat: Quick startup (created after setup)
- venv/: Virtual environment (created after setup)

Statistical Output Includes:
- Basic information and descriptive statistics
- ANOVA table with F-statistic and p-value
- Means comparison with confidence intervals
- Variance homogeneity tests
- Tukey HSD multiple comparisons
- Visual plots and charts

Support:
For questions about statistical interpretation, refer to standard ANOVA textbooks or statistical software documentation.