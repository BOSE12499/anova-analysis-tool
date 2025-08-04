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
from flask import Flask, request, jsonify, send_from_directory, make_response
from flask_cors import CORS

# Configure matplotlib for memory efficiency
matplotlib.rcParams['figure.max_open_warning'] = 0
matplotlib.rcParams['agg.path.chunksize'] = 10000  # Reduce path complexity
plt.ioff()  # Turn off interactive mode

# Try to import additional packages, set flags for availability
try:
    import pingouin as pg
    _PINGOUIN_AVAILABLE = True
except ImportError:
    _PINGOUIN_AVAILABLE = False
    print("Warning: pingouin not available. Some variance tests may use scipy fallbacks.")

try:
    from scipy.stats import studentized_range
    _STUDENTIZED_RANGE_AVAILABLE = True
except ImportError:
    print("Warning: studentized_range not available in your scipy version.")
    studentized_range = None
    _STUDENTIZED_RANGE_AVAILABLE = False

try:
    from statsmodels.stats.multicomp import MultiComparison
    _MULTICOMPARISON_AVAILABLE = True
except ImportError:
    print("Warning: statsmodels not available. Tukey HSD may not work.")
    MultiComparison = None
    _MULTICOMPARISON_AVAILABLE = False

# Initialize Flask app
app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}}) # Enable CORS for all routes

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
                }
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
        levene_dfnum, levene_dfden = np.nan, np.nan
        brown_forsythe_dfnum, brown_forsythe_dfden = np.nan, np.nan
        bartlett_dfnum = np.nan

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

                    bartlett_stat, bartlett_p_value = stats.bartlett(*groups_for_levene_scipy)
                    bartlett_dfnum = filtered_df_for_variance_test['LOT'].nunique() - 1
            else:
                levene_stat, levene_p_value = stats.levene(*groups_for_levene_scipy, center='mean')
                levene_dfnum = filtered_df_for_variance_test['LOT'].nunique() - 1
                levene_dfden = len(filtered_df_for_variance_test) - filtered_df_for_variance_test['LOT'].nunique()

                brown_forsythe_stat, brown_forsythe_p_value = stats.levene(*groups_for_levene_scipy, center='median')
                brown_forsythe_dfnum = levene_dfnum
                brown_forsythe_dfden = levene_dfden

                bartlett_stat, bartlett_p_value = stats.bartlett(*groups_for_levene_scipy)
                bartlett_dfnum = filtered_df_for_variance_test['LOT'].nunique() - 1

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
            'madStats': mad_stats_final,
            'plots': plots_base64
        }

        if tukey_results:
            response_data['tukey'] = tukey_results

        return jsonify(response_data)

    except Exception as e:
        return jsonify({"error": str(e), "traceback": "Check server logs for detailed traceback."}), 500

@app.route('/')
def index():
    # ตรวจสอบไฟล์ที่มีอยู่จริง
    html_files = ['my.html', 'index.html', 'calculator.html']
    for html_file in html_files:
        if os.path.exists(html_file):
            return send_from_directory('.', html_file)
    return jsonify({"error": "HTML file not found"}), 404

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

# เพิ่ม route สำหรับ health check
@app.route('/health')
def health_check():
    return jsonify({"status": "OK", "message": "Server is running"})

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    host = '0.0.0.0'  # สำหรับ production
    debug = os.environ.get('FLASK_ENV') == 'development'
    
    print(f"Starting server on host {host}, port {port}")
    print(f"Debug mode: {debug}")
    print(f"Available files: {[f for f in os.listdir('.') if f.endswith('.html')]}")
    app.run(host=host, port=port, debug=debug)