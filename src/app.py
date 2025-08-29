import pandas as pd
import numpy as np
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
from itertools import combinations

# Flask imports
from flask import Flask, request, jsonify, send_from_directory, make_response, render_template
from flask_cors import CORS

# Configure matplotlib for production deployment
matplotlib.rcParams['figure.max_open_warning'] = 0
matplotlib.rcParams['agg.path.chunksize'] = 10000
matplotlib.rcParams['figure.figsize'] = [6, 4]  # Smaller default figure size
matplotlib.rcParams['savefig.dpi'] = 60  # Lower DPI for production
plt.ioff()  # Turn off interactive mode

# Try to import additional packages with better error handling
try:
    import pingouin as pg
    _PINGOUIN_AVAILABLE = True
    print("Pingouin available")
except ImportError:
    _PINGOUIN_AVAILABLE = False
    print("WARNING: Pingouin not available. Using scipy fallbacks.")

try:
    from scipy.stats import studentized_range
    _STUDENTIZED_RANGE_AVAILABLE = True
    print("Studentized range available")
except ImportError:
    print("WARNING: Studentized range not available. Using chi2 approximation.")
    studentized_range = None
    _STUDENTIZED_RANGE_AVAILABLE = False

try:
    from statsmodels.stats.multicomp import MultiComparison
    _MULTICOMPARISON_AVAILABLE = True
    print("Statsmodels available")
except ImportError:
    print("WARNING: Statsmodels not available. Tukey HSD will be limited.")
    MultiComparison = None
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
        
        print(f"Bartlett Test (Excel Formula):")
        print(f"  Pooled Variance (sp): {sp:.6f}")
        print(f"  M statistic: {M:.6f}")
        print(f"  C correction: {C:.6f}")
        print(f"  Chi-square stat (M/C): {chi_square_stat:.6f}")
        print(f"  F Ratio (M/C)/(k-1): {bartlett_f_ratio:.6f}")
        print(f"  p-value: {bartlett_p_value:.6f}")
        
        return bartlett_f_ratio, bartlett_p_value, k-1
        
    except Exception as e:
        print(f"Warning: Bartlett Excel test failed: {e}")
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
        
        print(f"O'Brien[.5] Test (Excel Formula):")
        print(f"  Groups: {k}")
        print(f"  Group sizes: {group_ns}")
        print(f"  n (constant): {n_constant}")
        print(f"  Group means: {[f'{m:.6f}' for m in group_means]}")
        print(f"  Group variances: {[f'{v:.6f}' for v in group_vars]}")
        print(f"  Transformed means: {[f'{m:.6f}' for m in transformed_means]}")
        print(f"  Grand mean: {grand_mean:.6f}")
        print(f"  SSB = {n_constant} * {sum((mean_t - grand_mean)**2 for mean_t in transformed_means):.6f} = {SSB:.6f}")
        print(f"  SSW: {SSW:.6f}")
        print(f"  df_between: {df_between}, df_within: {df_within}")
        print(f"  MSB = SSB/{df_between} = {MSB:.6f}")
        print(f"  MSW = SSW/{df_within} = {MSW:.6f}")
        print(f"  F-statistic = MSB/MSW = {f_stat:.6f}")
        print(f"  p-value: {p_value:.6f}")
        
        return f_stat, p_value, df_between, df_within
        
    except Exception as e:
        print(f"Warning: O'Brien Excel test failed: {e}")
        return np.nan, np.nan, np.nan, np.nan

def plot_to_base64(plt):
    """Memory-optimized plot conversion with aggressive cleanup"""
    buf = io.BytesIO()
    try:
        # ลด DPI สำหรับ free tier และ optimize สำหรับ web
        plt.savefig(buf, format='png', bbox_inches='tight', 
                    dpi=75,  # ลดจาก 150 เป็น 75 (ประหยัด memory 75%)
                    facecolor='white', edgecolor='none',
                    transparent=False, pad_inches=0.05)  # ลบ optimize=True ที่ทำให้ error
        buf.seek(0)
        img_str = base64.b64encode(buf.getvalue()).decode('utf-8')
        return img_str
    finally:
        # Aggressive memory cleanup
        plt.close('all')  # Close all matplotlib figures
        buf.close()
        plt.clf()  # Clear current figure
        plt.cla()  # Clear current axis
        gc.collect()  # Force garbage collection


@app.route('/analyze_anova', methods=['POST', 'OPTIONS'])
def analyze_anova():
    try:
        # รับข้อมูล JSON จาก request - เพิ่มการตรวจสอบ
        print(f"Request content type: {request.content_type}")
        print(f"Request data: {request.data}")
        
        # ตรวจสอบว่าเป็น JSON request หรือไม่
        if request.content_type != 'application/json':
            return jsonify({"error": "Content-Type must be application/json"}), 400
            
        data = request.json
        if data is None:
            return jsonify({"error": "Invalid JSON data received"}), 400
            
        print(f"Parsed JSON data: {data}")
        
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

        if df['DATA'].isnull().any():
             return jsonify({"error": "CSV data contains missing (NaN) values in the 'DATA' column. Please clean your data."}), 400

        # --- Degrees of Freedom ---
        df_between = k_groups - 1
        df_within = n_total - k_groups
        df_total = n_total - 1

        if df_within <= 0:
            return jsonify({"error": f"Degrees of Freedom for Error (df_within) must be greater than 0. (Total data points: {n_total}, Number of groups: {k_groups}). Not enough data for ANOVA analysis."}), 400

        # --- Grand Mean & Group Means ---
        grand_mean = df['DATA'].mean()
        group_means = df.groupby('LOT')['DATA'].mean().to_dict() # Convert to dict
        group_stds = df.groupby('LOT')['DATA'].std().to_dict() # For variance test plot

        # --- Sum of Squares ---
        ss_total = np.sum((df['DATA'] - grand_mean) ** 2)
        ss_between = 0
        for lot in group_means:
            n_group = lot_counts[lot]
            group_mean = group_means[lot]
            ss_between += n_group * (group_mean - grand_mean) ** 2

        ss_within = np.sum((df['DATA'] - df.groupby('LOT')['DATA'].transform('mean')) ** 2)

        # --- Mean Squares ---
        ms_between = ss_between / df_between if df_between > 0 else 0
        ms_within = ss_within / df_within if df_within > 0 else 0

        # --- F-statistic & p-value ---
        f_statistic = ms_between / ms_within if ms_within > 0 else 0
        p_value = 1 - stats.f.cdf(f_statistic, df_between, df_within)

        alpha = 0.05

        # --- Plots ---
        # --- Memory Optimized Plots Generation ---
        # Pre-clear all plots before starting
        plt.close('all')
        gc.collect()
        
        plots_base64 = {}

        # 1. Oneway Analysis (Box Plot) - ลดขนาดและ optimize
        plt.figure(figsize=(8, 5))  # ลดจาก (10, 6)
        df.boxplot(column='DATA', by='LOT', grid=False, widths=0.5, patch_artist=True,
                    boxprops=dict(facecolor='lightblue', color='black'),
                    medianprops=dict(color='red'),
                    showfliers=True)
        plt.scatter(range(1, len(group_means) + 1), [group_means[lot] for lot in sorted(group_means.keys())],
                    color='green', marker='o', s=60, zorder=5, label='Group Means')  # ลด marker size

        if lsl is not None:
            plt.axhline(y=lsl, color='red', linestyle='--', linewidth=1.5, label='LSL')
        if usl is not None:
            plt.axhline(y=usl, color='red', linestyle='--', linewidth=1.5, label='USL')

        plt.title("Oneway Analysis of DATA by LOT")
        plt.suptitle("")
        plt.xlabel("LOT")
        plt.ylabel("DATA")
        plt.legend()
        plt.tight_layout()
        plots_base64['onewayAnalysisPlot'] = plot_to_base64(plt)
        
        # Force cleanup after each plot
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
            mean_val = lot_data.mean()
            std_dev_val = lot_data.std()

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
        print(f"Debug Tukey: k_groups={k_groups}, df_within={df_within}")
        print(f"Debug Tukey: MultiComparison available: {_MULTICOMPARISON_AVAILABLE}")
        print(f"Debug Tukey: lot_names={lot_names}")
        
        # เงื่อนไขการทำ Tukey HSD
        if k_groups < 2:
            print("Debug: ไม่สามารถทำการทดสอบ Tukey-Kramer HSD ได้ เนื่องจากมี LOT น้อยกว่า 2 กลุ่ม")
        elif df_within <= 0:
            print("Debug: ไม่สามารถทำการทดสอบ Tukey-Kramer HSD ได้ เนื่องจาก Degrees of Freedom สำหรับ Error (df_within) ไม่เพียงพอ")
        elif not _MULTICOMPARISON_AVAILABLE or MultiComparison is None:
            print("Debug: ไม่สามารถทำการทดสอบ Tukey-Kramer HSD ได้ เนื่องจาก MultiComparison ไม่พร้อมใช้งาน")
        else:
            print("Debug: เริ่มคำนวณ Tukey-Kramer HSD...")
            try:
                # Tukey-Kramer HSD test
                mc = MultiComparison(df['DATA'], df['LOT'])
                tukey_result = mc.tukeyhsd(alpha=alpha)
                print("Debug: Tukey HSD calculation successful")

                # 1. Confidence Quantile (q*)
                if _STUDENTIZED_RANGE_AVAILABLE and studentized_range is not None:
                    q_crit = studentized_range.ppf(1 - alpha, k_groups, df_within)
                    print(f"Debug: Using studentized_range, q_crit={q_crit}")
                else:
                    # Fallback: ใช้ค่าประมาณจาก chi-square
                    from scipy.stats import chi2
                    q_crit = np.sqrt(2 * chi2.ppf(1 - alpha, k_groups - 1))
                    print(f"Debug: Using chi2 approximation, q_crit={q_crit}")
                
                q_crit_for_jmp_display = q_crit / math.sqrt(2)

                # 2. --- HSD Threshold Matrix ---
                print("\n" + "="*80)
                print("                                HSD THRESHOLD MATRIX")
                print("="*80)

                # Create HSD threshold matrix exactly like EDIT.py
                lot_names = sorted(df['LOT'].unique())
                hsd_matrix = {}
                
                print("Abs(Dif)-HSD (Positive = Significant):")
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

                # Print HSD Matrix for debugging
                print("Debug: HSD Matrix created:")
                for lot_i in lot_names:
                    row_str = f"{lot_i}: "
                    for lot_j in lot_names:
                        value = hsd_matrix[lot_i][lot_j]
                        if value is None:
                            row_str += "    None    "
                        else:
                            row_str += f"{value:>10.6f} "
                    print(row_str)

                # 3. --- Connecting Letters Report ---
                from collections import defaultdict

                # Re-run Tukey HSD for clean summary table
                tukey_result_for_letters = MultiComparison(df['DATA'], df['LOT']).tukeyhsd(alpha=alpha)
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
                plt.figure(figsize=(8, min(5, len(ordered_diffs_df_sorted) * 0.4)))  # Dynamic size based on data
                y_pos_sorted = np.arange(len(ordered_diffs_df_sorted))
                differences_sorted = [d['rawDiff'] for d in ordered_diffs_df_sorted]
                lower_bounds_sorted = [d['lowerCL'] for d in ordered_diffs_df_sorted]
                upper_bounds_sorted = [d['upperCL'] for d in ordered_diffs_df_sorted]
                labels_sorted = [f"{d['lot1']} - {d['lot2']}" for d in ordered_diffs_df_sorted]

                lower_errors = [diff - lower for diff, lower in zip(differences_sorted, lower_bounds_sorted)]
                upper_errors = [upper - diff for diff, upper in zip(differences_sorted, upper_bounds_sorted)]

                plt.errorbar(differences_sorted, y_pos_sorted,
                                xerr=[lower_errors, upper_errors],
                                fmt='o', color='blue', ecolor='black', capsize=4, markersize=4)  # ลด marker size

                plt.axvline(x=0, linestyle='--', color='gray', linewidth=1)  # ลด line width
                plt.yticks(y_pos_sorted, labels_sorted, fontsize=9)  # ลด font size
                plt.xlabel("Mean Difference")
                plt.title("Tukey HSD Confidence Intervals")
                plt.grid(True, axis='x', linestyle='--', alpha=0.4)  # ลด alpha
                plt.tight_layout()
                plots_base64['tukeyChart'] = plot_to_base64(plt)
                
                # Final cleanup after Tukey chart
                gc.collect()

                tukey_results = {
                    'qCrit': q_crit_for_jmp_display,
                    'connectingLetters': connecting_letters_final,
                    'connectingLettersTable': connecting_letters_data,
                    'comparisons': ordered_diffs_df_sorted,
                    'hsdMatrix': hsd_matrix,  # Make sure this is included
                }
                
                print("Debug: HSD Matrix created:")
                print(f"Debug: HSD Matrix keys: {list(hsd_matrix.keys())}")
                print(f"Debug: Sample HSD Matrix values:")
                for i, (lot1, row) in enumerate(hsd_matrix.items()):
                    if i < 2:  # Show first 2 rows as sample
                        print(f"  {lot1}: {row}")
                print("Debug: Tukey results created successfully")
                print(f"Debug: tukey_results keys: {tukey_results.keys()}")
                
            except Exception as e:
                print(f"Error in Tukey HSD calculation: {str(e)}")
                import traceback
                print(f"Full traceback: {traceback.format_exc()}")
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

                except Exception as e:
                    print(f"Warning: Pingouin failed, falling back to scipy.stats for variance tests: {e}")
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

            # Plot Variance Chart - Memory optimized
            plt.figure(figsize=(7, 4))  # ลดจาก (8, 5)
            valid_group_stds = filtered_df_for_variance_test.groupby('LOT')['DATA'].std()
            lot_names_valid = sorted(valid_group_stds.index.tolist())
            std_dev_values_valid = [valid_group_stds[lot] for lot in lot_names_valid]

            plt.plot(lot_names_valid, std_dev_values_valid, 'o', color='black', markersize=5)  # ลด marker size
            plt.axhline(y=pooled_std, color='blue', linestyle=':', linewidth=1.2, 
                       label=f'Pooled Std Dev = {pooled_std:.4f}')  # ลด precision
            plt.xlabel("Lot")
            plt.ylabel("Std Dev")
            plt.title("Tests that the Variances are Equal")
            plt.ylim(bottom=0)
            plt.grid(axis='y', linestyle='--', alpha=0.5)  # ลด alpha
            plt.legend()
            plt.tight_layout()
            plots_base64['varianceChart'] = plot_to_base64(plt)
            
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
                welch_result = pg.welch_anova(data=df, dv='DATA', between='LOT')
                
                welch_results_data = {
                    'available': True,
                    'fStatistic': float(welch_result['F'].iloc[0]),
                    'dfNum': float(welch_result['ddof1'].iloc[0]),
                    'dfDen': float(welch_result['ddof2'].iloc[0]),
                    'pValue': float(welch_result['p-unc'].iloc[0])
                }
                
                print(f"Welch's ANOVA: F={welch_results_data['fStatistic']:.4f}, p={welch_results_data['pValue']:.4f}")
                
            except Exception as e:
                print(f"Error calculating Welch's ANOVA: {e}")
                welch_results_data = {'available': False, 'error': str(e)}
        else:
            welch_results_data = {'available': False, 'error': 'Pingouin not available'}

        # --- Mean Absolute Deviations ---
        mad_stats_final = []
        for lot in sorted(df['LOT'].unique()):
            lot_data = df[df['LOT'] == lot]['DATA']
            lot_count = len(lot_data)
            lot_std = lot_data.std()
            lot_mean = lot_data.mean()
            lot_median = lot_data.median()

            mad_to_mean = np.mean(np.abs(lot_data - lot_mean))
            mad_to_median = np.mean(np.abs(lot_data - lot_median))

            mad_stats_final.append({
                'Level': lot,
                'Count': lot_count,
                'Std Dev': lot_std if lot_count >= 2 else None,
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
        print("DEBUG: Attempting to render dashboard.html")
        return render_template('dashboard.html')
    except Exception as e:
        print(f"ERROR in dashboard(): {str(e)}")
        import traceback
        print(f"TRACEBACK: {traceback.format_exc()}")
        return jsonify({"error": str(e)}), 500

@app.route('/')
def index():
    # แสดงหน้า my.html เป็นหน้าหลัก
    try:
        print("DEBUG: Attempting to render my.html")
        return render_template('my.html')
    except Exception as e:
        print(f"ERROR in index(): {str(e)}")
        import traceback
        print(f"TRACEBACK: {traceback.format_exc()}")
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

if __name__ == '__main__':
    # Production configuration
    port = int(os.environ.get('PORT', 10000))
    host = '0.0.0.0'  # สำหรับ production deployment
    debug = os.environ.get('FLASK_ENV') != 'production'  # debug เฉพาะใน development
    
    print(f"🚀 Starting ANOVA Analysis Tool Server")
    print(f"📍 Host: {host}, Port: {port}")
    print(f"🐛 Debug mode: {debug}")
    print(f"🌍 Environment: {os.environ.get('FLASK_ENV', 'development')}")
    
    if debug:
        print(f"📱 Local URL: http://localhost:{port}")
        print(f"🌐 Network URL: http://{host}:{port}")
    
    print("="*50)
    app.run(host=host, port=port, debug=debug)