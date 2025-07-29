import pandas as pd
import numpy as np
from scipy import stats
import warnings
import math
from statsmodels.stats.multicomp import pairwise_tukeyhsd, MultiComparison
from scipy.stats import studentized_range
import matplotlib.pyplot as plt
import io
import base64
import json # เพิ่ม import สำหรับ json

# Flask imports
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS # เพิ่ม Flask-CORS สำหรับจัดการ Cross-Origin

# เพิ่ม import สำหรับ pingouin และตรวจสอบการติดตั้ง
try:
    import pingouin as pg
    _PINGOUIN_AVAILABLE = True
except ImportError:
    # ใน Production environment ไม่ควรใช้ pip install แบบนี้
    # ควรจัดการ dependencies ผ่าน requirements.txt ตั้งแต่ต้น
    _PINGOUIN_AVAILABLE = False
    warnings.warn("Pingouin library not found. Levene's and Bartlett's tests might fall back to scipy.stats.")

warnings.filterwarnings('ignore')

# Initialize Flask app
app = Flask(__name__)
CORS(app) # Enable CORS for all routes

def custom_round_up(value, decimals=5):
    """
    Custom rounding function to match JMP behavior (แบบโค้ดที่ 2)
    """
    multiplier = 10 ** decimals
    return np.ceil(value * multiplier) / multiplier

def plot_to_base64(plt):
    """Converts a matplotlib plot to a base64 encoded PNG string."""
    buf = io.BytesIO()
    plt.savefig(buf, format='png', bbox_inches='tight')
    plt.close() # Close the plot to free memory
    return base64.b64encode(buf.getvalue()).decode('utf-8')


@app.route('/analyze_anova', methods=['POST'])
def analyze_anova():
    try:
        # รับข้อมูล JSON จาก request
        data = request.json
        csv_data_string = data.get('csv_data')
        lsl = data.get('LSL')
        usl = data.get('USL')

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
        plots_base64 = {}

        # 1. Oneway Analysis (Box Plot)
        plt.figure(figsize=(10, 6))
        df.boxplot(column='DATA', by='LOT', grid=False, widths=0.5, patch_artist=True,
                    boxprops=dict(facecolor='lightblue', color='black'),
                    medianprops=dict(color='red'),
                    showfliers=True)
        plt.scatter(range(1, len(group_means) + 1), [group_means[lot] for lot in sorted(group_means.keys())],
                    color='green', marker='o', s=80, zorder=5, label='Group Means')

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
        if p_value < alpha and k_groups >= 2 and df_within > 0:
            mc = MultiComparison(df['DATA'], df['LOT'])
            tukey_result_obj = mc.tukeyhsd(alpha=alpha)

            q_crit = studentized_range.ppf(1 - alpha, k_groups, df_within)
            q_crit_for_jmp_display = q_crit / math.sqrt(2)

            # Connecting Letters Report
            # This is a simplified approach for connecting letters. JMP's is more sophisticated.
            # For perfect JMP matching, a more complex clustering logic would be needed.
            sorted_groups_by_mean = sorted(lot_names, key=lambda x: group_means[x], reverse=True)
            connecting_letters_map = {}
            current_letter = 'A'
            clusters = [] # List of lists, each sublist is a cluster of non-significantly different groups

            for g1 in sorted_groups_by_mean:
                found_cluster = False
                for cluster_idx, cluster in enumerate(clusters):
                    is_compatible = True
                    for g_in_cluster in cluster:
                        # Check if g1 is significantly different from any group in the current cluster
                        comp_row = tukey_result_obj.summary().data[1:] # Skip header
                        for row_data in comp_row:
                            group1_comp, group2_comp, _, _, p_adj_comp, reject_comp = row_data
                            if ((group1_comp == g1 and group2_comp == g_in_cluster) or
                                (group1_comp == g_in_cluster and group2_comp == g1)) and reject_comp:
                                is_compatible = False
                                break
                        if not is_compatible:
                            break
                    if is_compatible:
                        clusters[cluster_idx].append(g1)
                        found_cluster = True
                        break
                if not found_cluster:
                    clusters.append([g1]) # Create new cluster

            # Assign letters based on clusters
            temp_letter_assignments = {group: [] for group in lot_names}
            for i, cluster in enumerate(clusters):
                letter = chr(ord('A') + i)
                for group in cluster:
                    temp_letter_assignments[group].append(letter)

            for group in lot_names:
                temp_letter_assignments[group].sort()
                connecting_letters_map[group] = "".join(temp_letter_assignments[group])

            connecting_letters_data = []
            for g in sorted_groups_by_mean:
                count = lot_counts[g]
                mean_val = group_means[g]
                se_group = pooled_std / np.sqrt(count)
                connecting_letters_data.append({
                    'Level': g,
                    'Letter': connecting_letters_map.get(g, ''),
                    'Mean': mean_val,
                    'Std Error': se_group
                })

            # Ordered Differences Report
            ordered_diffs_data = []
            from itertools import combinations
            all_pairs = list(combinations(lot_names, 2))

            for lot_a, lot_b in all_pairs:
                mean_a = group_means[lot_a]
                mean_b = group_means[lot_b]
                ni, nj = lot_counts[lot_a], lot_counts[lot_b]

                std_err_diff_for_pair = np.sqrt(ms_within * (1/ni + 1/nj))
                margin_of_error_ci = q_crit * std_err_diff_for_pair / math.sqrt(2)

                diff_raw = mean_a - mean_b
                lower_cl_raw = diff_raw - margin_of_error_ci
                upper_cl_raw = diff_raw + margin_of_error_ci

                # Find p-adj from statsmodels output
                p_adj = np.nan
                for row_data in tukey_result_obj.summary().data[1:]:
                    group1_comp, group2_comp, _, _, p_adj_comp, _ = row_data
                    if (group1_comp == lot_a and group2_comp == lot_b) or \
                       (group1_comp == lot_b and group2_comp == lot_a):
                        p_adj = float(p_adj_comp)
                        break

                is_significant = p_adj < alpha if not np.isnan(p_adj) else False

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
            ).to_dict(orient='records') # Convert to list of dicts for JSON

            # Plot Tukey HSD Confidence Intervals
            plt.figure(figsize=(9, 6))
            y_pos = np.arange(len(ordered_diffs_data)) # Use original unsorted to match y_pos
            differences_plot = [d['rawDiff'] for d in ordered_diffs_data] # use raw diff for plot
            lower_bounds_plot = [d['lowerCL'] for d in ordered_diffs_data]
            upper_bounds_plot = [d['upperCL'] for d in ordered_diffs_data]
            labels_plot = [f"{d['lot1']} - {d['lot2']}" for d in ordered_diffs_data]

            # Re-sort for plotting order if needed, or adjust y_pos
            # For consistency with table, plot based on `ordered_diffs_df_sorted`
            y_pos_sorted = np.arange(len(ordered_diffs_df_sorted))
            differences_sorted = [d['rawDiff'] for d in ordered_diffs_df_sorted]
            lower_bounds_sorted = [d['lowerCL'] for d in ordered_diffs_df_sorted]
            upper_bounds_sorted = [d['upperCL'] for d in ordered_diffs_df_sorted]
            labels_sorted = [f"{d['lot1']} - {d['lot2']}" for d in ordered_diffs_df_sorted]

            lower_errors = [diff - lower for diff, lower in zip(differences_sorted, lower_bounds_sorted)]
            upper_errors = [upper - diff for diff, upper in zip(differences_sorted, upper_bounds_sorted)]

            plt.errorbar(differences_sorted, y_pos_sorted,
                            xerr=[lower_errors, upper_errors],
                            fmt='o', color='blue', ecolor='black', capsize=5)
            plt.axvline(x=0, linestyle='--', color='gray')
            plt.yticks(y_pos_sorted, labels_sorted)
            plt.xlabel("Mean Difference")
            plt.title("Tukey HSD Confidence Intervals (Ordered Differences)")
            plt.grid(True, axis='x', linestyle='--', alpha=0.6)
            plt.tight_layout()
            plots_base64['tukeyChart'] = plot_to_base64(plt)

            tukey_results = {
                'qCrit': q_crit_for_jmp_display,
                'connectingLetters': connecting_letters_map, # This is the simple map
                'connectingLettersTable': connecting_letters_data, # For display table in JS
                'comparisons': ordered_diffs_df_sorted, # Already sorted for display
            }

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
                    levene_results_pg = pg.homoscedasticity(data=filtered_df_for_variance_test, dv='DATA', group='LOT', method='levene', center='mean')
                    levene_stat = levene_results_pg['F'].iloc[0]
                    levene_p_value = levene_results_pg['p-unc'].iloc[0]
                    levene_dfnum = int(levene_results_pg['ddof1'].iloc[0])
                    levene_dfden = int(levene_results_pg['ddof2'].iloc[0])

                    brown_forsythe_results_pg = pg.homoscedasticity(data=filtered_df_for_variance_test, dv='DATA', group='LOT', method='levene', center='median')
                    brown_forsythe_stat = brown_forsythe_results_pg['F'].iloc[0]
                    brown_forsythe_p_value = brown_forsythe_results_pg['p-unc'].iloc[0]
                    brown_forsythe_dfnum = int(brown_forsythe_results_pg['ddof1'].iloc[0])
                    brown_forsythe_dfden = int(brown_forsythe_results_pg['ddof2'].iloc[0])

                    bartlett_results_pg = pg.homoscedasticity(data=filtered_df_for_variance_test, dv='DATA', group='LOT', method='bartlett')
                    bartlett_stat = bartlett_results_pg['W'].iloc[0]
                    bartlett_p_value = bartlett_results_pg['p-unc'].iloc[0]
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

            # Plot Variance Chart
            plt.figure(figsize=(8, 5))
            valid_group_stds = filtered_df_for_variance_test.groupby('LOT')['DATA'].std()
            lot_names_valid = sorted(valid_group_stds.index.tolist())
            std_dev_values_valid = [valid_group_stds[lot] for lot in lot_names_valid]

            plt.plot(lot_names_valid, std_dev_values_valid, 'o', color='black', markersize=6)
            plt.axhline(y=pooled_std, color='blue', linestyle=':', linewidth=1.5, label=f'Pooled Std Dev (RMSE) = {pooled_std:.5f}')
            plt.xlabel("Lot")
            plt.ylabel("Std Dev")
            plt.title("Tests that the Variances are Equal")
            plt.ylim(bottom=0)
            plt.grid(axis='y', linestyle='--', alpha=0.6)
            plt.legend()
            plt.tight_layout()
            plots_base64['varianceChart'] = plot_to_base64(plt)


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
    return send_from_directory('.', 'my.html')

@app.route('/<path:filename>')
def serve_static(filename):
    return send_from_directory('.', filename)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))  # Render ใช้ port 10000
    app.run(host='0.0.0.0', port=port, debug=False)