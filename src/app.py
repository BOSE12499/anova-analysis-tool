import pandas as pd
import numpy as np
# Set numpy precision to maximum for all calculations
np.set_printoptions(precision=15, suppress=False)
import scipy.stats as stats
import matplotlib
# Force matplotlib to use Agg backend before importing pyplot
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import seaborn as sns
import io
import base64
import json
import os
import math
import gc  # garbage collector for memory management
import threading
import time
from itertools import combinations
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

# Flask imports
from flask import Flask, request, jsonify, send_from_directory, make_response, render_template, session, send_file
from flask_cors import CORS
import logging
import warnings
import os

# ปิด warnings และ logging ที่ไม่จำเป็น
warnings.filterwarnings('ignore')
os.environ['OUTDATED_IGNORE'] = '1'  # ปิด outdated package warnings
logging.getLogger('werkzeug').setLevel(logging.ERROR)
logging.getLogger('flask').setLevel(logging.ERROR)

# Optimize matplotlib settings for performance
plt.rcParams.update({
    'figure.max_open_warning': 0,  # Disable warning about too many figures
    'font.size': 8,  # Smaller default font
    'axes.linewidth': 0.5,  # Thinner axes
    'lines.linewidth': 1.0,  # Thinner lines
})

# Global thread pool for async operations
_THREAD_POOL = ThreadPoolExecutor(max_workers=2)

# Simple cache for plot generation
_PLOT_CACHE = {}
_CACHE_MAX_SIZE = 50

def clear_plot_cache():
    """Clear plot cache to prevent memory buildup"""
    global _PLOT_CACHE
    if len(_PLOT_CACHE) > _CACHE_MAX_SIZE:
        # Keep only the most recent 25 entries
        items = list(_PLOT_CACHE.items())
        _PLOT_CACHE = dict(items[-25:])
        gc.collect()

def generate_cache_key(*args):
    """Generate a simple cache key from arguments"""
    return hash(str(args))

def async_plot_generation(plot_func, *args, **kwargs):
    """Generate plots asynchronously for better performance"""
    def _generate():
        return optimized_plot_to_base64(plot_func, *args, **kwargs)
    
    return _THREAD_POOL.submit(_generate)

# PowerPoint imports
try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.enum.text import PP_ALIGN
    from pptx.dml.color import RGBColor
    _PPTX_AVAILABLE = True
except ImportError:
    _PPTX_AVAILABLE = False

# Configure matplotlib for production deployment
matplotlib.rcParams['figure.max_open_warning'] = 0
matplotlib.rcParams['agg.path.chunksize'] = 10000
matplotlib.rcParams['figure.figsize'] = [6, 4]  # Smaller default figure size
matplotlib.rcParams['savefig.dpi'] = 60  # Lower DPI for production
plt.ioff()  # Turn off interactive mode

# Try to import additional packages with lazy loading for better performance
_PINGOUIN_AVAILABLE = None
_STUDENTIZED_RANGE_AVAILABLE = None
_MULTICOMPARISON_AVAILABLE = None

def get_pingouin():
    """Lazy loading for pingouin"""
    global _PINGOUIN_AVAILABLE
    if _PINGOUIN_AVAILABLE is None:
        try:
            import pingouin as pg
            _PINGOUIN_AVAILABLE = pg
        except ImportError:
            _PINGOUIN_AVAILABLE = False
    return _PINGOUIN_AVAILABLE

def get_studentized_range():
    """Lazy loading for studentized_range"""
    global _STUDENTIZED_RANGE_AVAILABLE
    if _STUDENTIZED_RANGE_AVAILABLE is None:
        try:
            from scipy.stats import studentized_range
            _STUDENTIZED_RANGE_AVAILABLE = studentized_range
        except ImportError:
            _STUDENTIZED_RANGE_AVAILABLE = False
    return _STUDENTIZED_RANGE_AVAILABLE

def get_multicomparison():
    """Lazy loading for MultiComparison"""
    global _MULTICOMPARISON_AVAILABLE
    if _MULTICOMPARISON_AVAILABLE is None:
        try:
            from statsmodels.stats.multicomp import MultiComparison
            _MULTICOMPARISON_AVAILABLE = MultiComparison
        except ImportError:
            _MULTICOMPARISON_AVAILABLE = False
    return _MULTICOMPARISON_AVAILABLE
    _MULTICOMPARISON_AVAILABLE = False

# Initialize Flask app with correct template folder
app = Flask(__name__, 
            template_folder='../templates')  # ระบุ path ไปยัง templates folder
CORS(app, resources={r"/*": {"origins": "*"}})

# Production configuration
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
app.config['JSON_SORT_KEYS'] = False
app.config['JSONIFY_PRETTYPRINT_REGULAR'] = False

# เพิ่ม OPTIONS handler สำหรับ preflight requests
@app.before_request
def handle_preflight():
    if request.method == "OPTIONS":
        response = make_response()
        response.headers.add("Access-Control-Allow-Origin", "*")
        response.headers.add('Access-Control-Allow-Headers', "*")
        response.headers.add('Access-Control-Allow-Methods', "*")
        return response

def custom_round_up(value, decimals=5):
    """
    Custom rounding function to match JMP behavior (แบบโค้ดที่ 2)
    """
    multiplier = 10 ** decimals
    return np.ceil(value * multiplier) / multiplier

def calculate_bartlett_excel(groups):
    """
    Bartlett test calculation based on Excel formula:
    =LET(
        d,D2:D31,e,E2:E31,f,F2:F31,g,G2:G31,
        vd,VAR(d),ve,VAR(e),vf,VAR(f),vg,VAR(g),
        sp,(29*vd+29*ve+29*vf+29*vg)/116,
        M,116*LN(sp)-29*LN(vd)-29*LN(ve)-29*LN(vf)-29*LN(vg),
        C,1+(1/9)*(4/29-1/116),
        (M/C)/3
    )
    """
    try:
        if len(groups) < 2:
            return np.nan, np.nan, np.nan
            
        # Calculate variances for each group (using Excel VAR function = sample variance)
        variances = []
        n_values = []
        
        for group in groups:
            if len(group) < 2:
                continue
            n = len(group)
            variance = np.var(group, ddof=1)  # Sample variance (n-1) like Excel VAR
            variances.append(variance)
            n_values.append(n)
        
        if len(variances) < 2:
            return np.nan, np.nan, np.nan
        
        k = len(variances)  # number of groups
        
        # Excel formula implementation:
        # sp = (29*vd + 29*ve + 29*vf + 29*vg) / 116
        # General form: sp = sum((ni-1)*vari) / sum(ni-1)
        numerator_sp = sum((n-1) * var for n, var in zip(n_values, variances))
        denominator_sp = sum(n-1 for n in n_values)  # This is total_n - k
        sp = numerator_sp / denominator_sp
        
        # M = 116*LN(sp) - 29*LN(vd) - 29*LN(ve) - 29*LN(vf) - 29*LN(vg)
        # General form: M = sum(ni-1)*ln(sp) - sum((ni-1)*ln(vari))
        M = denominator_sp * np.log(sp)
        for i, (n, var) in enumerate(zip(n_values, variances)):
            if var > 0:  # Avoid log(0)
                M -= (n-1) * np.log(var)
        
        # C = 1 + (1/9) * (4/29 - 1/116)
        # General form: C = 1 + (1/(3*(k-1))) * (sum(1/(ni-1)) - 1/sum(ni-1))
        sum_reciprocal = sum(1/(n-1) for n in n_values)
        reciprocal_total_df = 1/denominator_sp
        C = 1 + (1/(3*(k-1))) * (sum_reciprocal - reciprocal_total_df)
        
        # Excel formula final result: (M/C)/3
        # This gives F-ratio: (M/C)/(k-1)
        bartlett_f_ratio = (M / C) / (k - 1)
        
        # Calculate p-value using chi-square distribution
        chi_square_stat = M / C  # Traditional Bartlett statistic
        bartlett_p_value = 1 - stats.chi2.cdf(chi_square_stat, k-1)
        
        return bartlett_f_ratio, bartlett_p_value, k-1
        
    except Exception as e:
        return np.nan, np.nan, np.nan

def calculate_obrien_excel(groups):
    """
    O'Brien[.5] test calculation based on Excel formula:
    =LET(
        d,D2:D31,e,E2:E31,f,F2:F31,g,G2:G31,
        n,30,
        mean_d,AVERAGE(d),mean_e,AVERAGE(e),mean_f,AVERAGE(f),mean_g,AVERAGE(g),
        var_d,VAR(d),var_e,VAR(e),var_f,VAR(f),var_g,VAR(g),

        rd,MAP(d,LAMBDA(x,(n-1.5)*n*(x-mean_d)^2-0.5*var_d*(n-1))),
        re,MAP(e,LAMBDA(x,(n-1.5)*n*(x-mean_e)^2-0.5*var_e*(n-1))),
        rf,MAP(f,LAMBDA(x,(n-1.5)*n*(x-mean_f)^2-0.5*var_f*(n-1))),
        rg,MAP(g,LAMBDA(x,(n-1.5)*n*(x-mean_g)^2-0.5*var_g*(n-1))),

        mean_rd,AVERAGE(rd),mean_re,AVERAGE(re),mean_rf,AVERAGE(rf),mean_rg,AVERAGE(rg),
        grand_mean,AVERAGE(VSTACK(rd,re,rf,rg)),

        SSB,n*((mean_rd-grand_mean)^2+(mean_re-grand_mean)^2+(mean_rf-grand_mean)^2+(mean_rg-grand_mean)^2),
        SSW,SUMSQ(rd-mean_rd)+SUMSQ(re-mean_re)+SUMSQ(rf-mean_rf)+SUMSQ(rg-mean_rg),
        MSB,SSB/3,MSW,SSW/116,
        MSB/MSW
    )
    """
    try:
        if len(groups) < 2:
            return np.nan, np.nan, np.nan, np.nan
            
        # Store group data
        group_data = []
        group_means = []
        group_vars = []
        group_ns = []
        transformed_groups = []
        
        for group in groups:
            if len(group) < 2:
                continue
            n = len(group)
            mean_group = np.mean(group)
            var_group = np.var(group, ddof=1)  # Sample variance like Excel VAR
            
            group_data.append(group)
            group_means.append(mean_group)
            group_vars.append(var_group)
            group_ns.append(n)
            
            # O'Brien transformation: r = (n-1.5)*n*(x-mean)^2 - 0.5*var*(n-1)
            r_values = []
            for x in group:
                r = (n - 1.5) * n * (x - mean_group)**2 - 0.5 * var_group * (n - 1)
                r_values.append(r)
            
            transformed_groups.append(r_values)
        
        if len(transformed_groups) < 2:
            return np.nan, np.nan, np.nan, np.nan
        
        k = len(transformed_groups)  # number of groups
        
        # Calculate means of transformed values for each group
        transformed_means = [np.mean(group) for group in transformed_groups]
        
        # Calculate grand mean of all transformed values
        all_transformed = [val for group in transformed_groups for val in group]
        grand_mean = np.mean(all_transformed)
        
        # Calculate SSB (Sum of Squares Between) - Excel Formula
        # SSB = n*((mean_rd-grand_mean)^2+(mean_re-grand_mean)^2+(mean_rf-grand_mean)^2+(mean_rg-grand_mean)^2)
        # In Excel formula, n is constant for all groups (balanced design)
        # If groups have equal size, use that size; otherwise use smallest group size for consistency
        if len(set(group_ns)) == 1:
            # All groups have equal size (balanced design)
            n_constant = group_ns[0]
        else:
            # Unbalanced design - use common group size or smallest for conservative estimate
            n_constant = min(group_ns)
        
        SSB = n_constant * sum((mean_t - grand_mean)**2 for mean_t in transformed_means)
        
        # Calculate SSW (Sum of Squares Within)
        # SSW = SUMSQ(rd-mean_rd)+SUMSQ(re-mean_re)+...
        SSW = 0
        for i, (group_t, mean_t) in enumerate(zip(transformed_groups, transformed_means)):
            for val in group_t:
                SSW += (val - mean_t)**2
        
        # Degrees of freedom
        df_between = k - 1  # 3 in Excel example
        df_within = sum(group_ns) - k  # 116 in Excel example
        
        # Calculate Mean Squares
        MSB = SSB / df_between
        MSW = SSW / df_within
        
        # F-ratio
        f_stat = MSB / MSW
        
        # Calculate p-value
        p_value = 1 - stats.f.cdf(f_stat, df_between, df_within)
        
        return f_stat, p_value, df_between, df_within
        
    except Exception as e:
        return np.nan, np.nan, np.nan, np.nan

def optimized_plot_to_base64(plot_func, *args, **kwargs):
    """Enhanced plot conversion with professional styling and larger size"""
    # Pre-cleanup to ensure clean state
    plt.close('all')
    plt.clf()
    plt.cla()
    
    # Force garbage collection before creating new plot
    gc.collect()
    
    buf = io.BytesIO()
    fig = None
    try:
        # Create professional figure with slightly smaller size
        plt.rcParams['figure.max_open_warning'] = 0
        fig, ax = plt.subplots(figsize=(9, 5), dpi=100,  # Slightly reduced size
                              facecolor='white', edgecolor='none')
        
        # Enhanced styling
        plt.style.use('default')  # Clean base style
        fig.patch.set_facecolor('white')
        
        # Execute plot function with error handling
        try:
            plot_func(ax, *args, **kwargs)
        except Exception as e:
            # Fallback for any plotting errors
            ax.text(0.5, 0.5, f'Plot Error: {str(e)[:50]}...', 
                   ha='center', va='center', transform=ax.transAxes,
                   fontsize=12, color='red')
            ax.set_title('Plot Generation Error', fontsize=14, fontweight='bold')
        
        # Apply professional styling
        ax.grid(True, alpha=0.3, linestyle='-', linewidth=0.5)
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.spines['left'].set_linewidth(0.8)
        ax.spines['bottom'].set_linewidth(0.8)
        
        # Improved layout
        plt.tight_layout(pad=2.0)
        
        # Save with higher quality settings
        buf = io.BytesIO()
        fig.savefig(buf, format='png', bbox_inches='tight', 
                   dpi=100, facecolor='white', edgecolor='none',
                   transparent=False, pad_inches=0.2,
                   metadata=None)
        buf.seek(0)
        
        # Convert to base64
        img_bytes = buf.getvalue()
        img_str = base64.b64encode(img_bytes).decode('utf-8')
        
        return img_str
        
    except Exception as e:
        # Ultimate fallback
        return ""
    finally:
        # Ultra-aggressive cleanup
        if fig is not None:
            plt.close(fig)
        if buf:
            buf.close()
        plt.close('all')
        plt.clf()
        plt.cla()
        
        # Force immediate memory cleanup
        if 'img_bytes' in locals():
            del img_bytes
        gc.collect()

def create_boxplot(ax, df, group_means, lsl=None, usl=None):
    """Enhanced professional box plot creation with green connecting line"""
    # Create beautiful box plot with enhanced styling
    box_plot = df.boxplot(column='DATA', by='LOT', ax=ax, 
                         grid=False, widths=0.6, patch_artist=True, 
                         showfliers=True, return_type='dict')
    
    # Enhanced box styling
    colors = ['#E3F2FD', '#BBDEFB', '#90CAF9', '#64B5F6', '#42A5F5', '#2196F3']
    if 'DATA' in box_plot:
        boxes = box_plot['DATA']['boxes']
        for i, box in enumerate(boxes):
            box.set_facecolor(colors[i % len(colors)])
            box.set_alpha(0.7)
            box.set_linewidth(1.2)
            box.set_edgecolor('#1565C0')
    
    # Enhanced group means with diamond markers only
    lot_names = sorted(group_means.keys())
    x_positions = range(1, len(lot_names) + 1)
    means_values = [group_means[lot] for lot in lot_names]
    
    # Add diamond markers for group means (removed green connecting line)
    ax.scatter(x_positions, means_values,
              color='#4CAF50', marker='D', s=80, zorder=10, 
              alpha=0.9, edgecolors='white', linewidth=2,
              label='Group Means')
    
    # Enhanced specification limits
    if lsl is not None:
        ax.axhline(y=lsl, color='#F44336', linestyle='-', 
                  linewidth=2.5, alpha=0.8, label=f'LSL = {lsl}')
    if usl is not None:
        ax.axhline(y=usl, color='#F44336', linestyle='-', 
                  linewidth=2.5, alpha=0.8, label=f'USL = {usl}')
    
    # Professional styling with smaller fonts
    ax.set_title("Oneway Analysis of DATA by LOT", 
                fontsize=13, fontweight='bold', pad=15)  # Reduced from 16 to 13
    ax.set_xlabel("LOT", fontsize=11, fontweight='bold')  # Reduced from 14 to 11
    ax.set_ylabel("DATA", fontsize=11, fontweight='bold')  # Reduced from 14 to 11
    
    # Remove the automatic title from boxplot
    ax.figure.suptitle('')
    
    # Enhanced grid and styling
    ax.grid(True, alpha=0.3, linestyle='-', linewidth=0.5)
    ax.set_facecolor('#FAFAFA')
    
    # Rotate x-axis labels for better readability with smaller fonts
    plt.setp(ax.get_xticklabels(), rotation=45, ha='right', fontsize=10)  # Reduced from 12 to 10
    ax.tick_params(axis='y', labelsize=10)  # Reduced from 12 to 10
    
    # Add legend if there are spec limits with smaller font
    if lsl is not None or usl is not None:
        ax.legend(fontsize=9, loc='upper right',  # Reduced from 11 to 9
                 frameon=True, fancybox=True, shadow=True)
def create_tukey_plot(ax, tukey_data, group_means):
    """Enhanced professional Tukey HSD plot with confidence intervals"""
    # Extract data for plotting
    differences = []
    lower_bounds = []
    upper_bounds = []
    comparison_labels = []
    colors = []
    
    # Process tukey data to create confidence interval plot
    for key, data in tukey_data.items():
        differences.append(data['difference'])
        lower_bounds.append(data['lower'])
        upper_bounds.append(data['upper'])
        comparison_labels.append(key)
        
        # Color based on significance
        if data.get('significant', False):
            colors.append('#F44336')  # Red for significant
        else:
            colors.append('#4CAF50')  # Green for not significant
    
    if differences:
        y_positions = range(len(differences))
        lower_errors = [abs(diff - lower) for diff, lower in zip(differences, lower_bounds)]
        upper_errors = [abs(upper - diff) for diff, upper in zip(differences, upper_bounds)]
        
        # Create professional error bar plot
        for i, (diff, y_pos, color) in enumerate(zip(differences, y_positions, colors)):
            ax.errorbar(diff, y_pos, 
                       xerr=[[lower_errors[i]], [upper_errors[i]]],
                       fmt='o', color=color, ecolor=color, 
                       capsize=4, markersize=8, linewidth=2.5,
                       alpha=0.8, capthick=2)
        
        # Enhanced styling with smaller fonts
        ax.axvline(x=0, linestyle='--', color='gray', alpha=0.8, linewidth=2)
        ax.set_yticks(y_positions)
        ax.set_yticklabels(comparison_labels, fontsize=10)  # Reduced from 12 to 10
        ax.set_xlabel("Mean Difference", fontsize=11, fontweight='bold')  # Reduced from 14 to 11
        ax.set_title("Tukey HSD Confidence Intervals", 
                    fontsize=13, fontweight='bold', pad=15)  # Reduced from 16 to 13
        
        # Professional grid and background
        ax.grid(True, axis='x', alpha=0.3, linestyle='-', linewidth=0.5)
        ax.set_facecolor('#FAFAFA')
        
        # Add legend for significance
        from matplotlib.patches import Patch
        legend_elements = [
            Patch(facecolor='#F44336', alpha=0.8, label='Significant'),
            Patch(facecolor='#4CAF50', alpha=0.8, label='Not Significant')
        ]
        ax.legend(handles=legend_elements, fontsize=9,  # Reduced from 11 to 9
                 loc='upper right', frameon=True, fancybox=True, shadow=True)
        
    else:
        # Fallback to group means comparison
        groups = list(group_means.keys())
        means = [group_means[g] for g in groups]
        
        bars = ax.bar(range(len(groups)), means, 
                     color='#2196F3', alpha=0.7, width=0.6,
                     edgecolor='#1565C0', linewidth=1.5)
        
        ax.set_xticks(range(len(groups)))
        ax.set_xticklabels(groups, rotation=45, ha='right', fontsize=10)  # Reduced from 12 to 10
        ax.set_title("Group Means Comparison", fontsize=13, fontweight='bold', pad=15)  # Reduced from 16 to 13
        ax.set_ylabel("Group Means", fontsize=11, fontweight='bold')  # Reduced from 14 to 11
        ax.grid(True, alpha=0.3)
        ax.set_facecolor('#FAFAFA')
def create_variance_plot(ax, group_stds, equal_var_p_value):
    """Enhanced professional variance scatter plot with actual standard deviation from MAD table"""
    groups = list(group_stds.keys())
    std_devs = list(group_stds.values())  # Use actual Std Dev values from MAD table
    
    # Calculate pooled standard deviation with 15 decimal precision
    pooled_std = round(sum(std_devs) / len(std_devs), 15)
    
    # Rainbow colors progression (kept for reference but using black as requested)
    rainbow_colors = [
        '#FF0000',  # Red
        '#FF8000',  # Orange  
        '#FFFF00',  # Yellow
        '#80FF00',  # Light Green
        '#00FF00',  # Green
        '#00FF80',  # Cyan Green
        '#00FFFF',  # Cyan
        '#0080FF',  # Light Blue
        '#0000FF',  # Blue
        '#8000FF',  # Purple
        '#FF00FF',  # Magenta
        '#FF0080'   # Pink
    ]
    
    # Create professional scatter plot with black color
    for i, (group, std_dev) in enumerate(zip(groups, std_devs)):
        ax.scatter(i, std_dev, s=120, color='black', alpha=0.9, 
                  edgecolors='white', linewidth=2.0, zorder=5)
        
        # Add value labels for standard deviation with 7 decimal places
        ax.annotate(f'{std_dev:.7f}', (i, std_dev), 
                   xytext=(0, 10), textcoords='offset points',
                   ha='center', va='bottom', fontsize=9, fontweight='bold',
                   bbox=dict(boxstyle='round,pad=0.3', facecolor='white', alpha=0.8))
    
    # Add pooled standard deviation line (blue dashed line without label)
    ax.axhline(y=pooled_std, color='#0080FF', linestyle='--', 
              linewidth=2.0, alpha=0.8)
    
    # Enhanced styling with smaller fonts
    ax.set_xticks(range(len(groups)))
    ax.set_xticklabels(groups, fontsize=10, fontweight='bold')  # Reduced from 12 to 10
    ax.set_xlabel("Lot", fontsize=11, fontweight='bold')  # Reduced from 14 to 11
    ax.set_ylabel("Std Dev", fontsize=11, fontweight='bold')  # Reduced from 14 to 11
    
    # Set Y-axis with more compact scale
    min_std = min(std_devs)
    max_std = max(std_devs)
    y_range = max_std - min_std
    y_margin = y_range * 0.1  # 10% margin
    ax.set_ylim(max(0, min_std - y_margin), max_std + y_margin)
    
    # Set Y-axis ticks with 5 decimal places and rounded numbers
    from matplotlib.ticker import MaxNLocator
    ax.yaxis.set_major_locator(MaxNLocator(nbins=6, prune='both'))
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{round(x, 5):.5f}'))
    
    # Professional title with test result matching the image style - smaller font
    test_result = "Unequal" if equal_var_p_value < 0.05 else "Equal"
    ax.set_title(f"Standard Deviation Analysis - {test_result} Variances\n(Levene Test p={equal_var_p_value:.4f})", 
                fontsize=13, fontweight='bold', pad=12)  # Reduced from 16 to 13, pad from 15 to 12
    
    # Professional grid and background with reduced styling
    ax.grid(True, alpha=0.3, linestyle='-', linewidth=0.4)
    ax.set_facecolor('#FAFAFA')
    
    # Set y-axis to start from 0 for better visualization
    ax.set_ylim(bottom=0)


@app.route('/analyze_anova', methods=['POST', 'OPTIONS'])
def analyze_anova():
    try:
        # รับข้อมูล JSON จาก request - เพิ่มการตรวจสอบ
        
        # ตรวจสอบว่าเป็น JSON request หรือไม่
        if request.content_type != 'application/json':
            return jsonify({"error": "Content-Type must be application/json"}), 400
            
        data = request.json
        if data is None:
            return jsonify({"error": "Invalid JSON data received"}), 400
            
        # รับข้อมูลจาก request - เพิ่มการตรวจสอบ None
        csv_data_string = data.get('csv_data') if data else None
        lsl = data.get('LSL') if data else None
        usl = data.get('USL') if data else None

        if not csv_data_string:
            return jsonify({"error": "No CSV data provided."}), 400

        # แปลง csv_data_string เป็น DataFrame
        df = pd.read_csv(io.StringIO(csv_data_string), header=None, usecols=[0, 1])
        df.columns = ['LOT', 'DATA']

        # ตรวจสอบว่าคอลัมน์ 'DATA' เป็นตัวเลข
        if not pd.api.types.is_numeric_dtype(df['DATA']):
            return jsonify({"error": "Column 'DATA' (second column) is not numeric. Please check your CSV file."}), 400

        # --- Basic Information ---
        n_total = len(df)
        k_groups = df['LOT'].nunique()
        lot_names = sorted(df['LOT'].unique().tolist()) # Convert to list for JSON
        lot_counts = df['LOT'].value_counts().sort_index().to_dict() # Convert to dict for JSON

        # --- Pre-calculate ALL group statistics ONCE with 15 decimal precision ---
        group_stats = df.groupby('LOT').agg({
            'DATA': ['count', 'mean', 'std', 'var', 'min', 'max']
        }).round(15)  # Use 15 decimal places for internal calculations
        group_stats.columns = ['count', 'mean', 'std', 'var', 'min', 'max']
        
        # Convert to optimized dictionaries with high precision
        group_means = group_stats['mean'].to_dict()
        group_stds = group_stats['std'].to_dict()
        group_variances = group_stats['var'].to_dict()

        if df['DATA'].isnull().any():
             return jsonify({"error": "CSV data contains missing (NaN) values in the 'DATA' column. Please clean your data."}), 400

        # --- Degrees of Freedom ---
        df_between = k_groups - 1
        df_within = n_total - k_groups
        df_total = n_total - 1

        if df_within <= 0:
            return jsonify({"error": f"Degrees of Freedom for Error (df_within) must be greater than 0. (Total data points: {n_total}, Number of groups: {k_groups}). Not enough data for ANOVA analysis."}), 400

        # --- Grand Mean (calculated with 15 decimal precision) ---
        grand_mean = round(df['DATA'].mean(), 15)

        # --- Sum of Squares (calculated with 15 decimal precision) ---
        ss_total = round(np.sum((df['DATA'] - grand_mean) ** 2), 15)
        ss_between = 0
        for lot in lot_counts:  # Use pre-calculated lot_counts and group_means
            n_group = lot_counts[lot]
            group_mean = group_means[lot]  # Use pre-calculated group means
            ss_between += n_group * (group_mean - grand_mean) ** 2
        ss_between = round(ss_between, 15)

        ss_within = round(np.sum((df['DATA'] - df.groupby('LOT')['DATA'].transform('mean')) ** 2), 15)

        # --- Mean Squares (calculated with 15 decimal precision) ---
        ms_between = round(ss_between / df_between, 15) if df_between > 0 else 0
        ms_within = round(ss_within / df_within, 15) if df_within > 0 else 0

        # --- F-statistic & p-value (calculated with 15 decimal precision) ---
        f_statistic = round(ms_between / ms_within, 15) if ms_within > 0 else 0
        p_value = round(1 - stats.f.cdf(f_statistic, df_between, df_within), 15)

        alpha = 0.05

        # --- Pre-calculate ALL group statistics ONCE for maximum efficiency ---
        group_stats = df.groupby('LOT').agg({
            'DATA': ['count', 'mean', 'std', 'var', 'min', 'max']
        }).round(6)
        group_stats.columns = ['count', 'mean', 'std', 'var', 'min', 'max']
        
        # Convert to optimized dictionaries
        lot_counts = group_stats['count'].to_dict()
        group_means = group_stats['mean'].to_dict()
        group_stds = group_stats['std'].to_dict()
        group_variances = group_stats['var'].to_dict()

        # --- Plots ---
        # --- Ultra-Optimized Sequential Plot Generation ---
        # Pre-clear all plots before starting
        plt.close('all')
        gc.collect()
        
        plots_base64 = {}

        # 1. Ultra-optimized sequential plot generation with minimal memory usage
        # Clear cache if needed
        clear_plot_cache()
        
        # Generate box plot with optimized function
        plots_base64['onewayAnalysisPlot'] = optimized_plot_to_base64(
            create_boxplot, df, group_means, lsl, usl
        )
        
        # Immediate cleanup after each plot
        gc.collect()

        # --- ANOVA Table ---
        anova_results = {
            'Source': ['Lot', 'Error', 'C Total'],
            'DF': [df_between, df_within, df_total],
            'Sum of Squares': [ss_between, ss_within, ss_total],
            'Mean Square': [ms_between, ms_within, np.nan],
            'F Ratio': [f_statistic, np.nan, np.nan],
            'Prob > F': [p_value, np.nan, np.nan]
        }

        # --- Means for Oneway Anova ---
        group_stats_data = []
        pooled_std = np.sqrt(ms_within)
        t_critical_pooled_se = stats.t.ppf(1 - alpha/2, df_within)

        for lot in sorted(group_means.keys()):
            count = lot_counts[lot]
            mean_val = group_means[lot]
            std_error = pooled_std / np.sqrt(count)
            lower_95_pooled = mean_val - t_critical_pooled_se * std_error
            upper_95_pooled = mean_val + t_critical_pooled_se * std_error
            group_stats_data.append({
                'Level': lot,
                'Number': count,
                'Mean': mean_val,
                'Std Error': std_error,
                'Lower 95%': lower_95_pooled,
                'Upper 95%': upper_95_pooled
            })

        # --- Means and Std Deviations (Individual) ---
        means_std_devs_data = []
        for lot in sorted(group_means.keys()):
            lot_data = df[df['LOT'] == lot]['DATA']
            count = len(lot_data)
            # Calculate using STDEV.S equivalent (pandas default std() with ddof=1)
            mean_val = lot_data.mean()
            std_dev_val = lot_data.std()  # This is equivalent to STDEV.S in Excel

            individual_se = np.nan
            individual_lower = np.nan
            individual_upper = np.nan

            if count >= 2:
                individual_se = std_dev_val / np.sqrt(count)
                individual_df_ind = count - 1
                if individual_df_ind > 0:
                    individual_t_critical = stats.t.ppf(1 - alpha/2, individual_df_ind)
                    individual_lower = mean_val - individual_t_critical * individual_se
                    individual_upper = mean_val + individual_t_critical * individual_se

            means_std_devs_data.append({
                'Level': lot,
                'Number': count,
                'Mean': mean_val,
                'Std Dev': std_dev_val if count >= 2 else np.nan,
                'Std Err': individual_se,
                'Lower 95%': individual_lower,
                'Upper 95%': individual_upper
            })

        # --- Tukey-Kramer HSD ---
        tukey_results = None
        
        # เงื่อนไขการทำ Tukey HSD with lazy loading
        multicomp = get_multicomparison()
        if k_groups >= 2 and df_within > 0 and multicomp:
            try:
                # Tukey-Kramer HSD test
                mc = multicomp(df['DATA'], df['LOT'])
                tukey_result = mc.tukeyhsd(alpha=alpha)

                # 1. Confidence Quantile (q*)
                studentized_range = get_studentized_range()
                if studentized_range:
                    q_crit = studentized_range.ppf(1 - alpha, k_groups, df_within)
                else:
                    # Fallback: ใช้ค่าประมาณจาก chi-square
                    from scipy.stats import chi2
                    q_crit = np.sqrt(2 * chi2.ppf(1 - alpha, k_groups - 1))
                
                q_crit_for_jmp_display = q_crit / math.sqrt(2)

                # 2. --- HSD Threshold Matrix ---
                # Create HSD threshold matrix exactly like EDIT.py
                lot_names = sorted(df['LOT'].unique())
                hsd_matrix = {}
                
                for lot_i in lot_names:
                    hsd_matrix[lot_i] = {}
                    for lot_j in lot_names:
                        if lot_i == lot_j:
                            # Diagonal: ABS(mean_i - mean_i) - HSD = 0 - HSD = -HSD (negative value)
                            ni = lot_counts[lot_i]
                            # HSD threshold for same group (self-comparison)
                            hsd_threshold = (q_crit / math.sqrt(2)) * np.sqrt(ms_within * (1/ni + 1/ni))
                            # ABS(0) - HSD = -HSD (negative value)
                            diagonal_value = 0 - hsd_threshold
                            hsd_matrix[lot_i][lot_j] = round(diagonal_value, 8)
                        else:
                            # Calculate like EDIT.py: Abs(Dif) - HSD
                            mean_diff = abs(group_means[lot_i] - group_means[lot_j])
                            ni, nj = lot_counts[lot_i], lot_counts[lot_j]
                            
                            # HSD threshold calculation (using q_crit/sqrt(2) like EDIT.py)
                            hsd_threshold = (q_crit / math.sqrt(2)) * np.sqrt(ms_within * (1/ni + 1/nj))
                            
                            # Abs(Dif) - HSD (positive = significant)
                            abs_dif_minus_hsd = mean_diff - hsd_threshold
                            hsd_matrix[lot_i][lot_j] = round(abs_dif_minus_hsd, 8)

                # 3. --- Connecting Letters Report ---
                from collections import defaultdict

                # Re-run Tukey HSD for clean summary table
                tukey_result_for_letters = multicomp(df['DATA'], df['LOT']).tukeyhsd(alpha=alpha)
                summary = tukey_result_for_letters.summary()
                reject_table = pd.DataFrame(data=summary.data[1:], columns=summary.data[0])
                reject_table['reject'] = reject_table['reject'].astype(bool)

                groups_for_letters = sorted(df['LOT'].unique())
                sorted_groups_by_mean = sorted(groups_for_letters, key=lambda x: group_means[x], reverse=True)

                # Initialize letters for each group - แก้ไขให้เป็น dict แทน list
                group_letters_map = {group: [] for group in groups_for_letters}
                clusters = [] # Each cluster is a list of groups that are not significantly different

                for g1 in sorted_groups_by_mean:
                    assigned = False
                    # Try to assign g1 to an existing cluster
                    for cluster_idx, cluster in enumerate(clusters):
                        is_compatible = True
                        for g_in_cluster in cluster:
                            # Check if g1 is significantly different from any group in the current cluster
                            comp_row = reject_table[
                                ((reject_table['group1'] == g1) & (reject_table['group2'] == g_in_cluster)) |
                                ((reject_table['group1'] == g_in_cluster) & (reject_table['group2'] == g1))
                            ]
                            if not comp_row.empty and comp_row['reject'].values[0]: # If significantly different
                                is_compatible = False
                                break
                        if is_compatible:
                            clusters[cluster_idx].append(g1)
                            assigned = True
                            break

                    if not assigned:
                        # Create a new cluster for g1
                        clusters.append([g1])

                # Assign letters based on clusters
                letter_mapping = {}
                for i, cluster in enumerate(clusters):
                    letter = chr(ord('A') + i)
                    for group in cluster:
                        if group not in letter_mapping:
                            letter_mapping[group] = []
                        letter_mapping[group].append(letter)

                # Sort letters for each group and convert to string
                connecting_letters_final = {}
                for group in letter_mapping:
                    letter_mapping[group].sort()
                    connecting_letters_final[group] = "".join(letter_mapping[group])

                connecting_letters_data = []
                n_counts = df.groupby('LOT')['DATA'].count()
                se_groups = pooled_std / np.sqrt(n_counts)

                for g in sorted_groups_by_mean:
                    letters = connecting_letters_final.get(g, '') # Get assigned letters
                    connecting_letters_data.append({
                        'Level': g,
                        'Letter': letters,
                        'Mean': group_means[g],
                        'Std Error': se_groups[g]
                    })

                # 4. --- Ordered Differences Report ---
                ordered_diffs_data = []

                # Get raw p-values from statsmodels
                tukey_df_raw_pvalues = pd.DataFrame(data=tukey_result._results_table.data[1:], columns=tukey_result._results_table.data[0])
                tukey_df_raw_pvalues['p-adj'] = tukey_df_raw_pvalues['p-adj'].astype(float)

                # Generate all unique pairs
                from itertools import combinations
                all_pairs = list(combinations(lot_names, 2))

                for lot_a, lot_b in all_pairs:
                    mean_a = group_means[lot_a]
                    mean_b = group_means[lot_b]

                    ni, nj = lot_counts[lot_a], lot_counts[lot_b]

                    std_err_diff_for_pair = np.sqrt(ms_within * (1/ni + 1/nj))

                    # Margin of error for Tukey-Kramer CI
                    margin_of_error_ci = q_crit * std_err_diff_for_pair / math.sqrt(2)

                    diff_raw = mean_a - mean_b

                    lower_cl_raw = diff_raw - margin_of_error_ci
                    upper_cl_raw = diff_raw + margin_of_error_ci

                    # Find the p-adjusted value for the current pair
                    p_adj_row = tukey_df_raw_pvalues[
                        ((tukey_df_raw_pvalues['group1'] == lot_a) & (tukey_df_raw_pvalues['group2'] == lot_b)) |
                        ((tukey_df_raw_pvalues['group1'] == lot_b) & (tukey_df_raw_pvalues['group2'] == lot_a))
                    ]
                    p_adj = p_adj_row['p-adj'].iloc[0] if not p_adj_row.empty else np.nan

                    # Adjust display order to always show positive difference (Larger Mean - Smaller Mean)
                    if diff_raw < 0:
                        display_level_a, display_level_b = lot_b, lot_a
                        display_diff = -diff_raw
                        display_lower_cl = -upper_cl_raw
                        display_upper_cl = -lower_cl_raw
                    else:
                        display_level_a, display_level_b = lot_a, lot_b
                        display_diff = diff_raw
                        display_lower_cl = lower_cl_raw
                        display_upper_cl = upper_cl_raw

                    is_significant = p_adj < alpha if not np.isnan(p_adj) else False

                    ordered_diffs_data.append({
                        'lot1': display_level_a,
                        'lot2': display_level_b,
                        'rawDiff': display_diff,
                        'stdErrDiff': std_err_diff_for_pair,
                        'lowerCL': display_lower_cl,
                        'upperCL': display_upper_cl,
                        'p_adj': p_adj,
                        'isSignificant': is_significant
                    })

                # Sort by Difference (desc), then Level (asc), then - Level (asc) 
                ordered_diffs_df_sorted = pd.DataFrame(ordered_diffs_data).sort_values(
                    by=['rawDiff', 'lot1', 'lot2'], ascending=[False, True, True]
                ).to_dict(orient='records')

                # Plot Tukey HSD Confidence Intervals - Memory optimized
                # Generate Tukey HSD plot using optimized function
                tukey_data = {}
                for comparison in ordered_diffs_df_sorted:
                    key = f"{comparison['lot1']}-{comparison['lot2']}"
                    tukey_data[key] = {
                        'significant': comparison['rawDiff'] < comparison['lowerCL'] or comparison['rawDiff'] > comparison['upperCL'],
                        'difference': comparison['rawDiff'],
                        'lower': comparison['lowerCL'],
                        'upper': comparison['upperCL']
                    }
                
                plots_base64['tukeyChart'] = optimized_plot_to_base64(
                    create_tukey_plot, tukey_data, group_means
                )
                
                # Final cleanup after Tukey chart
                gc.collect()

                tukey_results = {
                    'qCrit': q_crit_for_jmp_display,
                    'connectingLetters': connecting_letters_final,
                    'connectingLettersTable': connecting_letters_data,
                    'comparisons': ordered_diffs_df_sorted,
                    'hsdMatrix': hsd_matrix,  # Make sure this is included
                }
                
            except Exception as e:
                tukey_results = None

        # --- Tests that the Variances are Equal ---
        levene_stat, levene_p_value = np.nan, np.nan
        brown_forsythe_stat, brown_forsythe_p_value = np.nan, np.nan
        bartlett_stat, bartlett_p_value = np.nan, np.nan
        obrien_stat, obrien_p_value = np.nan, np.nan  # Add O'Brien variables
        levene_dfnum, levene_dfden = np.nan, np.nan
        brown_forsythe_dfnum, brown_forsythe_dfden = np.nan, np.nan
        bartlett_dfnum = np.nan
        obrien_dfnum, obrien_dfden = np.nan, np.nan  # Add O'Brien df variables

        filtered_df_for_variance_test = df.groupby('LOT').filter(lambda x: len(x) >= 2)

        if filtered_df_for_variance_test['LOT'].nunique() >= 2:
            groups_for_levene_scipy = [filtered_df_for_variance_test[filtered_df_for_variance_test['LOT'] == lot]['DATA'].values for lot in sorted(filtered_df_for_variance_test['LOT'].unique())]

            if _PINGOUIN_AVAILABLE:
                try:
                    # แก้ไขการใช้ pingouin โดยการแปลงข้อมูลให้ถูกต้อง
                    variance_test_df = filtered_df_for_variance_test.copy()
                    variance_test_df['LOT'] = variance_test_df['LOT'].astype(str)
                    variance_test_df['DATA'] = pd.to_numeric(variance_test_df['DATA'], errors='coerce')
                    
                    # ลบ NaN values ที่อาจเหลืออยู่
                    variance_test_df = variance_test_df.dropna()
                    
                    pg = get_pingouin()
                    if pg:
                        levene_results_pg = pg.homoscedasticity(data=variance_test_df, dv='DATA', group='LOT', method='levene', center='mean')
                        levene_stat = float(levene_results_pg['F'].iloc[0])
                        levene_p_value = float(levene_results_pg['p-unc'].iloc[0])
                        levene_dfnum = int(levene_results_pg['ddof1'].iloc[0])
                        levene_dfden = int(levene_results_pg['ddof2'].iloc[0])

                        brown_forsythe_results_pg = pg.homoscedasticity(data=variance_test_df, dv='DATA', group='LOT', method='levene', center='median')
                        brown_forsythe_stat = float(brown_forsythe_results_pg['F'].iloc[0])
                        brown_forsythe_p_value = float(brown_forsythe_results_pg['p-unc'].iloc[0])
                        brown_forsythe_dfnum = int(brown_forsythe_results_pg['ddof1'].iloc[0])
                        brown_forsythe_dfden = int(brown_forsythe_results_pg['ddof2'].iloc[0])

                        bartlett_results_pg = pg.homoscedasticity(data=variance_test_df, dv='DATA', group='LOT', method='bartlett')
                        bartlett_stat = float(bartlett_results_pg['W'].iloc[0])
                        bartlett_p_value = float(bartlett_results_pg['p-unc'].iloc[0])
                        bartlett_dfnum = int(bartlett_results_pg['ddof1'].iloc[0])
                    else:
                        raise ImportError("Pingouin not available")

                except Exception as e:
                    levene_stat, levene_p_value = stats.levene(*groups_for_levene_scipy, center='mean')
                    levene_dfnum = filtered_df_for_variance_test['LOT'].nunique() - 1
                    levene_dfden = len(filtered_df_for_variance_test) - filtered_df_for_variance_test['LOT'].nunique()

                    brown_forsythe_stat, brown_forsythe_p_value = stats.levene(*groups_for_levene_scipy, center='median')
                    brown_forsythe_dfnum = levene_dfnum
                    brown_forsythe_dfden = levene_dfden

                    # Use Excel-compatible Bartlett test
                    bartlett_stat, bartlett_p_value, bartlett_dfnum = calculate_bartlett_excel(groups_for_levene_scipy)
                    
                    # Use Excel-compatible O'Brien[.5] test
                    obrien_stat, obrien_p_value, obrien_dfnum, obrien_dfden = calculate_obrien_excel(groups_for_levene_scipy)
            else:
                levene_stat, levene_p_value = stats.levene(*groups_for_levene_scipy, center='mean')
                levene_dfnum = filtered_df_for_variance_test['LOT'].nunique() - 1
                levene_dfden = len(filtered_df_for_variance_test) - filtered_df_for_variance_test['LOT'].nunique()

                brown_forsythe_stat, brown_forsythe_p_value = stats.levene(*groups_for_levene_scipy, center='median')
                brown_forsythe_dfnum = levene_dfnum
                brown_forsythe_dfden = levene_dfden

                # Use Excel-compatible Bartlett test
                bartlett_stat, bartlett_p_value, bartlett_dfnum = calculate_bartlett_excel(groups_for_levene_scipy)
                bartlett_dfnum = filtered_df_for_variance_test['LOT'].nunique() - 1
                
                # Use Excel-compatible O'Brien[.5] test
                obrien_stat, obrien_p_value, obrien_dfnum, obrien_dfden = calculate_obrien_excel(groups_for_levene_scipy)

            # Calculate Std Dev using same method as tables (STDEV.S equivalent)
            chart_std_devs = {}
            for lot in sorted(df['LOT'].unique()):
                lot_data = df[df['LOT'] == lot]['DATA']
                if len(lot_data) >= 2:
                    chart_std_devs[lot] = lot_data.std()  # STDEV.S equivalent
                else:
                    chart_std_devs[lot] = np.nan

            # Generate Variance Chart using STDEV.S calculated Std Dev values
            plots_base64['varianceChart'] = optimized_plot_to_base64(
                create_variance_plot, chart_std_devs, levene_p_value
            )
            
            # Cleanup after variance chart
            gc.collect()


        levene_results_data = {
            'fStatistic': levene_stat,
            'pValue': levene_p_value,
            'dfNum': levene_dfnum,
            'dfDen': levene_dfden
        }
        brown_forsythe_results_data = {
            'fStatistic': brown_forsythe_stat,
            'pValue': brown_forsythe_p_value,
            'dfNum': brown_forsythe_dfnum,
            'dfDen': brown_forsythe_dfden
        }
        bartlett_results_data = {
            'statistic': bartlett_stat, # Renamed to statistic as it's Chi2, not F
            'pValue': bartlett_p_value,
            'dfNum': bartlett_dfnum
        }
        obrien_results_data = {
            'fStatistic': obrien_stat,  # O'Brien uses F-statistic
            'pValue': obrien_p_value,
            'dfNum': obrien_dfnum,
            'dfDen': obrien_dfden
        }


        # --- Welch's ANOVA (for unequal variances) ---
        welch_results_data = None
        if _PINGOUIN_AVAILABLE:
            try:
                # Perform Welch's ANOVA using Pingouin
                pg = get_pingouin()
                if pg:
                    welch_result = pg.welch_anova(data=df, dv='DATA', between='LOT')
                else:
                    # Fallback if pingouin not available
                    welch_result = None
                
                welch_results_data = {
                    'available': True,
                    'fStatistic': float(welch_result['F'].iloc[0]),
                    'dfNum': float(welch_result['ddof1'].iloc[0]),
                    'dfDen': float(welch_result['ddof2'].iloc[0]),
                    'pValue': float(welch_result['p-unc'].iloc[0])
                }
                
            except Exception as e:
                welch_results_data = {'available': False, 'error': str(e)}
        else:
            welch_results_data = {'available': False, 'error': 'Pingouin not available'}

        # --- Mean Absolute Deviations ---
        mad_stats_final = []
        for lot in sorted(df['LOT'].unique()):
            lot_data = df[df['LOT'] == lot]['DATA']
            lot_count = len(lot_data)
            
            # Use the same calculation method as Means and Std Deviations table (STDEV.S equivalent)
            lot_std = round(lot_data.std(), 15) if lot_count >= 2 else None  # STDEV.S equivalent
            lot_mean = round(lot_data.mean(), 15)
            lot_median = round(lot_data.median(), 15)

            mad_to_mean = round(np.mean(np.abs(lot_data - lot_mean)), 15)
            mad_to_median = round(np.mean(np.abs(lot_data - lot_median)), 15)

            mad_stats_final.append({
                'Level': lot,
                'Count': lot_count,
                'Std Dev': lot_std,
                'MeanAbsDif to Mean': mad_to_mean,
                'MeanAbsDif to Median': mad_to_median
            })

        # Final JSON Response
        response_data = {
            'basicInfo': {
                'totalPoints': n_total,
                'numLots': k_groups,
                'lotNames': lot_names,
                'groupCounts': lot_counts
            },
            'means': {
                'grandMean': grand_mean,
                'groupMeans': group_means,
                'groupStatsPooledSE': group_stats_data,
                'groupStatsIndividual': means_std_devs_data
            },
            'anova': {
                'dfBetween': df_between,
                'dfWithin': df_within,
                'dfTotal': df_total,
                'ssBetween': ss_between,
                'ssWithin': ss_within,
                'ssTotal': ss_total,
                'msBetween': ms_between,
                'msWithin': ms_within,
                'fStatistic': f_statistic,
                'pValue': p_value
            },
            'levene': levene_results_data,
            'brownForsythe': brown_forsythe_results_data,
            'bartlett': bartlett_results_data,
            'obrien': obrien_results_data,  # Add O'Brien test results
            'welch': welch_results_data,
            'madStats': mad_stats_final,
            'plots': plots_base64
        }

        if tukey_results:
            response_data['tukey'] = tukey_results

        return jsonify(response_data)

    except Exception as e:
        return jsonify({"error": str(e), "traceback": "Check server logs for detailed traceback."}), 500

@app.route('/dashboard')
def dashboard():
    """Serve the dashboard page"""
    try:
        return render_template('dashboard.html')
    except Exception as e:
        import traceback
        return jsonify({"error": str(e)}), 500

@app.route('/')
def index():
    # แสดงหน้า my.html เป็นหน้าหลัก
    try:
        return render_template('my.html')
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/<path:filename>')
def serve_static(filename):
    try:
        # ตรวจสอบว่าไฟล์มีอยู่จริง
        if os.path.exists(filename):
            return send_from_directory('.', filename)
        else:
            return jsonify({"error": f"File {filename} not found"}), 404
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# เพิ่ม route สำหรับ version information
@app.route('/version')
def get_version():
    try:
        # ลองหา VERSION.txt ในหลายตำแหน่ง
        version_paths = [
            '../docs/VERSION.txt',  # สำหรับ production
            'VERSION.txt',  # สำหรับ local development
            'docs/VERSION.txt'  # alternative path
        ]
        
        version_file_path = None
        for path in version_paths:
            if os.path.exists(path):
                version_file_path = path
                break
        
        if version_file_path:
            with open(version_file_path, 'r', encoding='utf-8') as f:
                version_content = f.read()
            
            # Extract version from content
            lines = version_content.split('\n')
            version = "v1.0.3"  # default version
            
            # ค้นหา version จากเนื้อหา
            for line in lines:
                if line.startswith('v') and '(' in line:
                    version = line.split(' ')[0]
                    break
            
            return jsonify({
                "version": version,
                "content": version_content,
                "status": "OK",
                "path": version_file_path
            })
        else:
            return jsonify({
                "version": "v1.0.3",
                "content": "Version file not found",
                "status": "File not found"
            })
    except Exception as e:
        return jsonify({
            "version": "v1.0.0",
            "error": str(e),
            "status": "Error"
        }), 500

# เพิ่ม route สำหรับ health check
@app.route('/health')
def health_check():
    return jsonify({"status": "OK", "message": "Server is running"})

def create_powerpoint_report(data, result, charts_data=None):
    """สร้างรายงาน PowerPoint เฉพาะผลการวิเคราะห์"""
    if not _PPTX_AVAILABLE:
        raise ImportError("python-pptx is not available")
    
    prs = Presentation()
    
    # Slide 1: Title Slide
    slide_layout = prs.slide_layouts[0]  # Title slide layout
    slide = prs.slides.add_slide(slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]
    
    title.text = "ANOVA Analysis Results"
    subtitle.text = f"Statistical Analysis Report\nGenerated on {datetime.now().strftime('%B %d, %Y')}"
    
    # Slide 2: ANOVA Results Table
    slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(slide_layout)
    title = slide.shapes.title
    title.text = "ANOVA Test Results"
    
    # Create table for ANOVA results
    rows, cols = 4, 6
    left = Inches(0.5)
    top = Inches(1.8)
    width = Inches(9)
    height = Inches(3.5)
    
    table = slide.shapes.add_table(rows, cols, left, top, width, height).table
    
    # Headers
    headers = ['Source of Variation', 'df', 'Sum of Squares', 'Mean Square', 'F-statistic', 'p-value']
    for i, header in enumerate(headers):
        cell = table.cell(0, i)
        cell.text = header
        cell.text_frame.paragraphs[0].font.bold = True
        cell.text_frame.paragraphs[0].font.size = Pt(11)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(54, 96, 146)
        cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
        cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    
    # ANOVA data
    anova_data = [
        ['Between Groups', f"{result['df_between']}", f"{result['ss_between']:.6f}", 
         f"{result['ms_between']:.6f}", f"{result['f_statistic']:.6f}", f"{result['p_value']:.6f}"],
        ['Within Groups (Error)', f"{result['df_within']}", f"{result['ss_within']:.6f}", 
         f"{result['ms_within']:.6f}", "-", "-"],
        ['Total', f"{result['df_total']}", f"{result['ss_total']:.6f}", "-", "-", "-"]
    ]
    
    for row_idx, row_data in enumerate(anova_data, 1):
        for col_idx, cell_data in enumerate(row_data):
            cell = table.cell(row_idx, col_idx)
            cell.text = str(cell_data)
            cell.text_frame.paragraphs[0].font.size = Pt(10)
            cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
            
            # Highlight significant p-value
            if col_idx == 5 and row_idx == 1 and cell_data != "-":
                try:
                    if float(cell_data) < 0.05:
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = RGBColor(255, 230, 230)
                except ValueError:
                    pass
    
    # Slide 3: Statistical Summary
    slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(slide_layout)
    title = slide.shapes.title
    title.text = "Statistical Summary"
    
    alpha = 0.05
    significant = result['p_value'] < alpha
    
    # Calculate effect size (eta squared) with 15 decimal precision
    eta_squared = round(result['ss_between'] / result['ss_total'], 15)
    
    # Effect size interpretation
    if eta_squared < 0.01:
        effect_size_text = "Very Small"
    elif eta_squared < 0.06:
        effect_size_text = "Small"
    elif eta_squared < 0.14:
        effect_size_text = "Medium"
    else:
        effect_size_text = "Large"
    
    summary_text = f"""Test Statistics:
• F-statistic: {result['f_statistic']:.6f}
• p-value: {result['p_value']:.6f}
• Degrees of freedom: {result['df_between']}, {result['df_within']}
• Effect size (η²): {eta_squared:.6f} ({effect_size_text})

Hypothesis Test (α = {alpha}):
• H₀: All group means are equal
• H₁: At least one group mean differs

Result: {'REJECT H₀' if significant else 'FAIL TO REJECT H₀'}

Conclusion:
{'There IS a statistically significant difference between group means.' if significant else 'There is NO statistically significant difference between group means.'}

Confidence Level: {(1-alpha)*100}%
Statistical Power: {'High' if significant and eta_squared > 0.06 else 'Moderate' if significant else 'Low'}"""

    content = slide.placeholders[1]
    content.text = summary_text
    content.text_frame.paragraphs[0].font.size = Pt(14)
    
    # Slide 4: Group Means Analysis
    slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(slide_layout)
    title = slide.shapes.title
    title.text = "Group Means Comparison"
    
    # Automatically detect column names and calculate group statistics
    try:
        # Try to identify group and value columns
        if len(data.columns) >= 2:
            group_col = data.columns[0]
            value_col = data.columns[1]
        else:
            # Fallback for single column or empty data
            return prs
        
        # Calculate group statistics
        group_stats = data.groupby(group_col)[value_col].agg(['count', 'mean', 'std'])
        
        # Create table for group means
        rows, cols = len(group_stats) + 1, 4
        table = slide.shapes.add_table(rows, cols, Inches(2), Inches(2), Inches(6), Inches(4)).table
        
        # Headers
        desc_headers = ['Group', 'N', 'Mean', 'Std. Deviation']
        for i, header in enumerate(desc_headers):
            cell = table.cell(0, i)
            cell.text = header
            cell.text_frame.paragraphs[0].font.bold = True
            cell.text_frame.paragraphs[0].font.size = Pt(12)
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(54, 96, 146)
            cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
            cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        
        # Data
        for row_idx, (group, stats) in enumerate(group_stats.iterrows(), 1):
            table.cell(row_idx, 0).text = str(group)
            table.cell(row_idx, 1).text = str(int(stats['count']))
            table.cell(row_idx, 2).text = f"{stats['mean']:.4f}"
            table.cell(row_idx, 3).text = f"{stats['std']:.4f}" if not pd.isna(stats['std']) else "N/A"
            
            # Center align all cells
            for col_idx in range(4):
                table.cell(row_idx, col_idx).text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
                table.cell(row_idx, col_idx).text_frame.paragraphs[0].font.size = Pt(11)
        
    except Exception as e:
        # Add error message to slide instead of table
        content = slide.placeholders[1]
        content.text = f"Group statistics could not be calculated.\nUsing analysis results from ANOVA calculations.\n\nStatistical results are still valid and shown in previous slides."
    
    # Slide 5: Post-hoc Analysis Information (only if significant)
    if significant:
        slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        title.text = "Post-hoc Analysis Recommendation"
        
        # Try to get group information from data, fallback to generic recommendations
        try:
            if len(data.columns) >= 2:
                group_col = data.columns[0]
                unique_groups = data[group_col].unique()
                num_groups = len(unique_groups)
                possible_comparisons = num_groups * (num_groups - 1) // 2
                
                group_list = chr(10).join([f"• {group}" for group in sorted(unique_groups)])
            else:
                num_groups = 2  # Default assumption
                possible_comparisons = 1
                group_list = "• Groups from analysis"
                
        except Exception as e:
            num_groups = 2  # Default assumption
            possible_comparisons = 1
            group_list = "• Groups from analysis"
        
        posthoc_text = f"""Since the ANOVA test shows significant differences (p = {result['p_value']:.6f}), 
post-hoc tests are recommended to identify which specific groups differ from each other.

Recommended Post-hoc Tests:
• Tukey's HSD: For equal sample sizes and equal variances
• Games-Howell: For unequal sample sizes or unequal variances  
• Bonferroni: For conservative pairwise comparisons
• Scheffé: For complex contrasts

Number of possible pairwise comparisons: {possible_comparisons}

Groups to compare:
{group_list}

Note: Conduct post-hoc tests to determine the specific nature of the differences."""

        content = slide.placeholders[1]
        content.text = posthoc_text
        content.text_frame.paragraphs[0].font.size = Pt(13)
    
    # Slide 6: ANOVA Assumptions
    slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(slide_layout)
    title = slide.shapes.title
    title.text = "ANOVA Assumptions"
    
    assumptions_text = """ANOVA Validity Assumptions:

1. Independence of Observations
   • Each observation should be independent of others
   • Random sampling or random assignment recommended

2. Normality
   • Data within each group should be approximately normally distributed
   • Central Limit Theorem helps with larger sample sizes (n ≥ 30)

3. Homogeneity of Variances (Homoscedasticity)
   • Variances should be approximately equal across groups
   • Levene's test can be used to check this assumption

4. No Extreme Outliers
   • Extreme outliers can affect the validity of results
   • Check box plots and identify potential outliers

Recommendation: Verify these assumptions before interpreting results.
Consider non-parametric alternatives (Kruskal-Wallis) if assumptions are violated."""

    content = slide.placeholders[1]
    content.text = assumptions_text
    content.text_frame.paragraphs[0].font.size = Pt(12)
    
    return prs

@app.route('/export_powerpoint', methods=['POST'])
def export_powerpoint():
    """Export ANOVA results เป็นไฟล์ PowerPoint"""
    try:
        if not _PPTX_AVAILABLE:
            return jsonify({'error': 'PowerPoint export is not available. Please install python-pptx.'}), 500
        
        # Get data from request
        request_data = request.get_json()
        if not request_data:
            return jsonify({'error': 'No data provided'}), 400
        
        # Validate required data
        if 'result' not in request_data:
            return jsonify({'error': 'No analysis results provided'}), 400
        
        result = request_data['result']
        charts_data = request_data.get('charts_data', {})
        
        # Handle data - can be empty or reconstructed
        data_input = request_data.get('data', [])
        if data_input and len(data_input) > 0:
            try:
                data = pd.DataFrame(data_input)
            except Exception as e:
                # Create minimal data for PowerPoint
                data = pd.DataFrame({
                    'Group': ['A', 'B'],
                    'Value': [25.0, 26.0]
                })
        else:
            # Create minimal data for PowerPoint when no data provided
            data = pd.DataFrame({
                'Group': ['A', 'B'],
                'Value': [25.0, 26.0]
            })
        
        # Validate result structure
        required_result_keys = ['f_statistic', 'p_value', 'df_between', 'df_within', 'df_total',
                               'ss_between', 'ss_within', 'ss_total', 'ms_between', 'ms_within']
        
        for key in required_result_keys:
            if key not in result:
                return jsonify({'error': f'Missing result key: {key}'}), 400
        
        # Create PowerPoint presentation
        prs = create_powerpoint_report(data, result, charts_data)
        
        # Save to memory
        pptx_io = io.BytesIO()
        prs.save(pptx_io)
        pptx_io.seek(0)
        
        # Create filename with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"ANOVA_Analysis_Report_{timestamp}.pptx"
        
        return send_file(
            pptx_io,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )
        
    except Exception as e:
        import traceback
        return jsonify({'error': f'Failed to create PowerPoint: {str(e)}'}), 500

if __name__ == '__main__':
    # Production configuration
    port = int(os.environ.get('PORT', 10000))
    host = '0.0.0.0'  # สำหรับ production deployment
    debug = os.environ.get('FLASK_ENV') != 'production'  # debug เฉพาะใน development
    
    # แสดงเฉพาะ localhost URL
    print(f"� ANOVA Analysis Tool - http://localhost:{port}")
    
    app.run(host=host, port=port, debug=debug)