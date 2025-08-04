# ANOVA Analysis Tool

A comprehensive web-based ANOVA (Analysis of Variance) tool with Tukey-Kramer HSD post-hoc analysis, designed to provide JMP-like statistical output.

## Features

- **One-way ANOVA Analysis** with detailed statistical output
- **Tukey-Kramer HSD Post-hoc Tests** for multiple comparisons
- **Variance Equality Tests** (Levene's, Brown-Forsythe, Bartlett's)
- **Interactive Box Plots** with specification limits
- **Mean Absolute Deviation Analysis**
- **Professional Statistical Reports** matching JMP format

## Live Demo

🚀 **Deployed on Render:** [Your-App-URL-Here]

## Local Development

### Prerequisites
- Python 3.11+
- pip

### Installation

1. Clone the repository:
```bash
git clone <your-repo-url>
cd WEB Calculator
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Run the application:
```bash
python app.py
```

4. Open your browser to `http://localhost:10000`

## Usage

1. **Upload CSV File**: Your file should have two columns:
   - Column A: LOT (categorical grouping variable)
   - Column B: DATA (numerical values)
   - No headers required

2. **Sample CSV Format**:
```csv
LOT1,10.2
LOT1,11.5
LOT1,10.8
LOT2,12.1
LOT2,11.8
LOT2,12.5
LOT3,9.8
LOT3,9.5
LOT3,10.0
```

3. **Optional Specification Limits**: Enter LSL and USL values for visualization

4. **Run Analysis**: Click "Perform ANOVA Analysis" to get comprehensive results

## Output Sections

- **Basic Information**: Data summary and group counts
- **ANOVA Table**: F-statistics, p-values, and significance tests
- **Means Analysis**: Group means with confidence intervals
- **Tukey-Kramer HSD**: Multiple comparison analysis with connecting letters
- **Variance Tests**: Levene's, Brown-Forsythe, and Bartlett's tests
- **Visual Charts**: Box plots, variance charts, and Tukey confidence intervals

## Technology Stack

- **Backend**: Flask (Python)
- **Statistical Libraries**: SciPy, Statsmodels, Pingouin
- **Visualization**: Matplotlib, Seaborn
- **Frontend**: HTML5, CSS3, JavaScript
- **Deployment**: Render.com

## License

MIT License - feel free to use for academic and commercial purposes.

## Contributing

Pull requests are welcome! For major changes, please open an issue first.