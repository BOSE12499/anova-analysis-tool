# Statistics Analysis 📊

Professional web-based statistical analysis tool for ANOVA (Analysis of Variance) with interactive dashboard and comprehensive export capabilities.

## 🚀 Features

### 📈 Statistical Analysis
- **One-way ANOVA** - Complete analysis of variance
- **Tukey-Kramer HSD** - Multiple comparison testing
- **Levene's Test** - Homogeneity of variance testing
- **Descriptive Statistics** - Comprehensive group statistics

### 🎨 Modern Interface
- **Responsive Design** - Works on all devices
- **Dark/Light Themes** - User preference support
- **Interactive Dashboard** - Real-time data visualization
- **Professional UI** - Glassmorphism design with smooth animations

### 📊 Export Capabilities
- **PDF Reports** - Professional statistical reports
- **Excel Spreadsheets** - Detailed data and results
- **PowerPoint Presentations** - Ready-to-present slides
- **JSON Data** - Raw analysis results

### 📱 Input Methods
- **CSV File Upload** - Standard data format support
- **Copy & Paste** - Direct text input with validation
- **Real-time Validation** - Instant feedback on data quality

## 🛠 Technology Stack

- **Backend**: Python Flask
- **Frontend**: HTML5, CSS3, JavaScript
- **Charts**: Chart.js
- **Statistics**: SciPy, Statsmodels, Pingouin
- **Data Processing**: Pandas, NumPy
- **Export**: ReportLab (PDF), OpenPyXL (Excel), python-pptx (PowerPoint)

## 📋 Requirements

```txt
Flask==3.0.3
pandas==2.2.2
numpy==1.26.4
scipy==1.13.1
matplotlib==3.9.2
python-pptx==0.6.23
reportlab==4.2.2
openpyxl==3.1.5
gunicorn==23.0.0
```

## 🚀 Quick Start

### Local Development
```bash
# Clone the repository
git clone https://github.com/BOSE12499/anova-analysis-tool.git
cd anova-analysis-tool

# Install dependencies
pip install -r requirements.txt

# Run the application
cd src
python app.py
```

Visit `http://localhost:10000` in your browser.

### Production Deployment
This application is ready for deployment on platforms like:
- Render.com
- Heroku  
- Railway
- Vercel
- Any cloud platform supporting Python/Flask

## 📊 Data Format

Upload CSV files with the following format:
```csv
LOT,Value
A,10.5
A,11.2
A,10.8
B,12.1
B,12.5
B,12.3
C,9.8
C,9.5
C,9.9
```

## 🎯 Usage

1. **Upload Data**: Use CSV upload or copy-paste interface
2. **Analyze**: Click "Analyze Data" to perform statistical tests
3. **View Results**: Review comprehensive statistical output
4. **Dashboard**: Click "Dashboard" for visualizations
5. **Export**: Choose from PDF, Excel, PowerPoint, or JSON formats

## 📈 Statistical Output

- **Basic Information**: Sample sizes, group counts
- **ANOVA Results**: F-statistic, p-values, effect sizes
- **Group Means**: Descriptive statistics for each group
- **Tukey-Kramer HSD**: Pairwise comparisons with confidence intervals
- **Levene's Test**: Variance homogeneity assessment
- **Effect Size**: Eta-squared and omega-squared calculations

## 🎨 UI Features

- **Modern Glassmorphism Design**
- **Gradient Backgrounds and Effects**
- **Smooth Animations and Transitions**
- **Responsive Grid Layouts**
- **Professional Export Modals**
- **Interactive Charts and Visualizations**

## 🔧 Configuration

Environment variables for production:
- `PORT`: Server port (default: 10000)
- `FLASK_ENV`: Environment (production/development)
- `VERSION`: Application version

## 📝 License

This project is licensed under the MIT License.

## 👨‍💻 Developer

**BOSE12499**
- GitHub: [@BOSE12499](https://github.com/BOSE12499)

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Submit a pull request

## 🐛 Issues

Please report issues on the [GitHub Issues](https://github.com/BOSE12499/anova-analysis-tool/issues) page.

---

**🚀 Ready for production deployment with modern UI and comprehensive statistical analysis capabilities!**