"""
Export Formatter Module
Professional formatting for ANOVA analysis exports
รองรับการ Export หลายรูปแบบ พร้อมการจัดรูปแบบที่เป็นมืออาชีพ
"""

import pandas as pd
import numpy as np
from typing import Dict, List, Any, Optional, Tuple
from datetime import datetime
import io
import base64
import json


class ExportFormatter:
    """
    Class สำหรับจัดรูปแบบการ Export ให้เป็นมืออาชีพ
    """
    
    def __init__(self):
        self.export_timestamp = datetime.now()
        self.format_version = "2.0"
    
    def format_for_pdf_export(self, result: Dict[str, Any], validation_result: Dict[str, Any]) -> Dict[str, Any]:
        """
        Format data specifically for PDF export with professional structure
        
        Args:
            result: ANOVA analysis results
            validation_result: Validation results
            
        Returns:
            Formatted data ready for PDF generation
        """
        formatted_data = {
            'header': self._create_professional_header(),
            'executive_summary': self._create_executive_summary(result, validation_result),
            'methodology': self._create_methodology_section(result),
            'raw_data_summary': self._create_raw_data_summary(result),
            'anova_results': self._format_anova_table(result.get('anova', {})),
            'descriptive_statistics': self._format_descriptive_stats(result.get('means', {})),
            'post_hoc_tests': self._format_post_hoc_results(result),
            'assumptions_checks': self._format_assumptions_tests(result),
            'interpretation': self._create_interpretation_section(result, validation_result),
            'conclusions': self._create_conclusions_section(result, validation_result),
            'appendices': self._create_appendices(result, validation_result),
            'footer': self._create_professional_footer(validation_result)
        }
        
        return formatted_data
    
    def format_for_excel_export(self, result: Dict[str, Any], validation_result: Dict[str, Any]) -> Dict[str, pd.DataFrame]:
        """
        Format data for Excel export with multiple sheets
        
        Args:
            result: ANOVA analysis results
            validation_result: Validation results
            
        Returns:
            Dictionary of DataFrames for Excel sheets
        """
        excel_data = {}
        
        # Summary sheet
        excel_data['Summary'] = self._create_summary_dataframe(result, validation_result)
        
        # ANOVA table sheet
        excel_data['ANOVA_Table'] = self._create_anova_dataframe(result.get('anova', {}))
        
        # Descriptive statistics sheet
        excel_data['Descriptive_Stats'] = self._create_descriptive_dataframe(result.get('means', {}))
        
        # Post-hoc tests sheet
        if 'tukey' in result:
            excel_data['Tukey_HSD'] = self._create_tukey_dataframe(result['tukey'])
        
        # Assumptions tests sheet
        excel_data['Assumptions'] = self._create_assumptions_dataframe(result)
        
        # Raw data (if available)
        if 'basicInfo' in result and 'rawGroups' in result['basicInfo']:
            excel_data['Raw_Data'] = self._create_raw_data_dataframe(result['basicInfo']['rawGroups'])
        
        # Validation report
        excel_data['Validation_Report'] = self._create_validation_dataframe(validation_result)
        
        return excel_data
    
    def format_for_csv_export(self, result: Dict[str, Any], export_type: str = 'summary') -> str:
        """
        Format data for CSV export
        
        Args:
            result: ANOVA analysis results
            export_type: Type of CSV export ('summary', 'anova', 'means', 'tukey')
            
        Returns:
            CSV formatted string
        """
        if export_type == 'summary':
            return self._create_summary_csv(result)
        elif export_type == 'anova':
            return self._create_anova_csv(result.get('anova', {}))
        elif export_type == 'means':
            return self._create_means_csv(result.get('means', {}))
        elif export_type == 'tukey':
            return self._create_tukey_csv(result.get('tukey', {}))
        else:
            return self._create_summary_csv(result)
    
    def _create_professional_header(self) -> Dict[str, str]:
        """Create professional header information"""
        return {
            'title': 'One-Way ANOVA Statistical Analysis Report',
            'subtitle': 'Comprehensive Statistical Analysis and Results',
            'timestamp': self.export_timestamp.strftime('%B %d, %Y at %H:%M:%S'),
            'software': 'ANOVA Analysis Tool v2.0',
            'analyst': 'Statistical Analysis System',
            'report_id': f"ANOVA_{self.export_timestamp.strftime('%Y%m%d_%H%M%S')}"
        }
    
    def _create_executive_summary(self, result: Dict[str, Any], validation_result: Dict[str, Any]) -> Dict[str, Any]:
        """Create executive summary section"""
        summary = {
            'title': 'Executive Summary',
            'analysis_type': 'One-Way Analysis of Variance (ANOVA)',
            'purpose': 'To test for significant differences between group means',
            'data_quality': 'Valid' if validation_result.get('is_valid', False) else 'Issues Detected',
            'key_findings': [],
            'statistical_significance': 'Not Determined',
            'recommendation': 'See detailed results below'
        }
        
        if 'anova' in result:
            anova = result['anova']
            p_value = anova.get('pValue', 1)
            f_stat = anova.get('fStatistic', 0)
            
            if p_value < 0.05:
                summary['statistical_significance'] = f'Significant (p = {p_value:.4f})'
                summary['key_findings'].append(f'F-statistic: {f_stat:.4f}')
                summary['key_findings'].append('Groups show statistically significant differences')
                summary['recommendation'] = 'Proceed with post-hoc analysis to identify specific group differences'
            else:
                summary['statistical_significance'] = f'Not Significant (p = {p_value:.4f})'
                summary['key_findings'].append(f'F-statistic: {f_stat:.4f}')
                summary['key_findings'].append('No statistically significant differences between groups')
                summary['recommendation'] = 'Groups do not differ significantly at α = 0.05 level'
        
        return summary
    
    def _create_methodology_section(self, result: Dict[str, Any]) -> Dict[str, Any]:
        """Create methodology section"""
        methodology = {
            'title': 'Statistical Methodology',
            'test_type': 'One-Way Analysis of Variance (ANOVA)',
            'hypothesis': {
                'null': 'H₀: μ₁ = μ₂ = ... = μₖ (all group means are equal)',
                'alternative': 'H₁: At least one group mean differs from the others'
            },
            'significance_level': 'α = 0.05',
            'test_statistic': 'F-ratio = MS_between / MS_within',
            'assumptions': [
                'Independence of observations',
                'Normality of residuals',
                'Homogeneity of variance (homoscedasticity)'
            ]
        }
        
        # Add sample size information if available
        if 'basicInfo' in result:
            basic_info = result['basicInfo']
            methodology['sample_info'] = {
                'total_observations': basic_info.get('totalN', 'N/A'),
                'number_of_groups': basic_info.get('groupCount', 'N/A'),
                'groups': basic_info.get('lotNames', [])
            }
        
        return methodology
    
    def _create_raw_data_summary(self, result: Dict[str, Any]) -> Dict[str, Any]:
        """Create raw data summary section"""
        summary = {
            'title': 'Data Summary',
            'overview': 'No data available'
        }
        
        if 'basicInfo' in result:
            basic_info = result['basicInfo']
            summary['overview'] = {
                'total_observations': basic_info.get('totalN', 'N/A'),
                'number_of_groups': basic_info.get('groupCount', 'N/A'),
                'group_names': basic_info.get('lotNames', []),
                'group_sizes': basic_info.get('groupCounts', {})
            }
        
        return summary
    
    def _format_anova_table(self, anova: Dict[str, Any]) -> Dict[str, Any]:
        """Format ANOVA table for professional display"""
        if not anova:
            return {'title': 'ANOVA Table', 'data': 'No ANOVA results available'}
        
        table_data = {
            'title': 'Analysis of Variance Table',
            'headers': ['Source', 'DF', 'Sum of Squares', 'Mean Square', 'F-Ratio', 'P-Value'],
            'rows': []
        }
        
        # Between groups row
        between_row = [
            'Between Groups',
            str(anova.get('dfBetween', 'N/A')),
            f"{anova.get('ssBetween', 0):.6f}",
            f"{anova.get('msBetween', 0):.6f}",
            f"{anova.get('fStatistic', 0):.6f}",
            f"{anova.get('pValue', 0):.6f}"
        ]
        table_data['rows'].append(between_row)
        
        # Within groups row
        within_row = [
            'Within Groups',
            str(anova.get('dfWithin', 'N/A')),
            f"{anova.get('ssWithin', 0):.6f}",
            f"{anova.get('msWithin', 0):.6f}",
            '—',  # No F-ratio for within groups
            '—'   # No p-value for within groups
        ]
        table_data['rows'].append(within_row)
        
        # Total row
        total_row = [
            'Total',
            str(anova.get('dfTotal', 'N/A')),
            f"{anova.get('ssTotal', 0):.6f}",
            '—',  # No mean square for total
            '—',  # No F-ratio for total
            '—'   # No p-value for total
        ]
        table_data['rows'].append(total_row)
        
        return table_data
    
    def _format_descriptive_stats(self, means: Dict[str, Any]) -> Dict[str, Any]:
        """Format descriptive statistics for professional display"""
        if not means or 'groupMeans' not in means:
            return {'title': 'Descriptive Statistics', 'data': 'No descriptive statistics available'}
        
        stats_data = {
            'title': 'Descriptive Statistics by Group',
            'headers': ['Group', 'N', 'Mean', 'Std Dev', 'Std Error', '95% CI Lower', '95% CI Upper'],
            'rows': []
        }
        
        group_means = means['groupMeans']
        group_stds = means.get('groupStds', {})
        group_counts = means.get('groupCounts', {})
        group_se = means.get('groupSE', {})
        group_ci_lower = means.get('groupCILower', {})
        group_ci_upper = means.get('groupCIUpper', {})
        
        for group in sorted(group_means.keys()):
            row = [
                str(group),
                str(group_counts.get(group, 'N/A')),
                f"{group_means[group]:.6f}",
                f"{group_stds.get(group, 0):.6f}",
                f"{group_se.get(group, 0):.6f}",
                f"{group_ci_lower.get(group, 0):.6f}",
                f"{group_ci_upper.get(group, 0):.6f}"
            ]
            stats_data['rows'].append(row)
        
        return stats_data
    
    def _format_post_hoc_results(self, result: Dict[str, Any]) -> Dict[str, Any]:
        """Format post-hoc test results"""
        if 'tukey' not in result:
            return {'title': 'Post-Hoc Tests', 'data': 'No post-hoc tests performed'}
        
        tukey = result['tukey']
        
        post_hoc_data = {
            'title': 'Tukey HSD Post-Hoc Test Results',
            'critical_value': tukey.get('criticalValue', 'N/A'),
            'confidence_level': '95%',
            'headers': ['Comparison', 'Mean Difference', 'P-Value', 'Significant', 'CI Lower', 'CI Upper'],
            'rows': []
        }
        
        if 'pairwiseComparisons' in tukey:
            comparisons = tukey['pairwiseComparisons']
            for pair, data in comparisons.items():
                row = [
                    pair,
                    f"{data.get('meanDiff', 0):.6f}",
                    f"{data.get('pValue', 0):.6f}",
                    'Yes' if data.get('significant', False) else 'No',
                    f"{data.get('ciLower', 0):.6f}",
                    f"{data.get('ciUpper', 0):.6f}"
                ]
                post_hoc_data['rows'].append(row)
        
        return post_hoc_data
    
    def _format_assumptions_tests(self, result: Dict[str, Any]) -> Dict[str, Any]:
        """Format statistical assumptions test results"""
        assumptions_data = {
            'title': 'Statistical Assumptions Tests',
            'tests': []
        }
        
        # Levene's test for homogeneity of variance
        if 'levene' in result:
            levene = result['levene']
            assumptions_data['tests'].append({
                'name': "Levene's Test for Homogeneity of Variance",
                'statistic': f"{levene.get('statistic', 0):.6f}",
                'p_value': f"{levene.get('pValue', 0):.6f}",
                'interpretation': 'Homogeneity assumed' if levene.get('pValue', 0) > 0.05 else 'Heterogeneity detected'
            })
        
        # Shapiro-Wilk test for normality
        if 'shapiro' in result:
            shapiro = result['shapiro']
            assumptions_data['tests'].append({
                'name': 'Shapiro-Wilk Test for Normality',
                'statistic': f"{shapiro.get('statistic', 0):.6f}",
                'p_value': f"{shapiro.get('pValue', 0):.6f}",
                'interpretation': 'Normality assumed' if shapiro.get('pValue', 0) > 0.05 else 'Non-normality detected'
            })
        
        if not assumptions_data['tests']:
            assumptions_data['tests'].append({
                'name': 'No assumption tests performed',
                'note': 'Standard ANOVA assumptions should be verified separately'
            })
        
        return assumptions_data
    
    def _create_interpretation_section(self, result: Dict[str, Any], validation_result: Dict[str, Any]) -> Dict[str, Any]:
        """Create statistical interpretation section"""
        interpretation = {
            'title': 'Statistical Interpretation',
            'overall_result': 'Not Available',
            'effect_size': 'Not Calculated',
            'practical_significance': 'Cannot be determined',
            'recommendations': []
        }
        
        if 'anova' in result:
            anova = result['anova']
            p_value = anova.get('pValue', 1)
            f_stat = anova.get('fStatistic', 0)
            
            if p_value < 0.05:
                interpretation['overall_result'] = f'Statistically significant difference between groups (F = {f_stat:.4f}, p = {p_value:.4f})'
                interpretation['recommendations'].extend([
                    'Reject the null hypothesis',
                    'At least one group mean differs significantly from the others',
                    'Consider post-hoc tests to identify specific group differences',
                    'Examine practical significance of observed differences'
                ])
            else:
                interpretation['overall_result'] = f'No statistically significant difference between groups (F = {f_stat:.4f}, p = {p_value:.4f})'
                interpretation['recommendations'].extend([
                    'Fail to reject the null hypothesis',
                    'No evidence of differences between group means',
                    'Consider power analysis if differences were expected',
                    'Review data collection and measurement procedures'
                ])
        
        # Add validation-related interpretations
        if validation_result.get('warnings'):
            interpretation['data_quality_notes'] = validation_result['warnings']
        
        return interpretation
    
    def _create_conclusions_section(self, result: Dict[str, Any], validation_result: Dict[str, Any]) -> Dict[str, Any]:
        """Create conclusions section"""
        conclusions = {
            'title': 'Conclusions and Recommendations',
            'primary_conclusion': 'Analysis incomplete',
            'statistical_conclusion': 'Not determined',
            'methodological_notes': [],
            'future_directions': []
        }
        
        if 'anova' in result and validation_result.get('is_valid', False):
            anova = result['anova']
            p_value = anova.get('pValue', 1)
            
            if p_value < 0.05:
                conclusions['primary_conclusion'] = 'Significant differences exist between the analyzed groups'
                conclusions['statistical_conclusion'] = f'The analysis provides strong evidence (p = {p_value:.4f}) against the null hypothesis of equal group means'
                conclusions['future_directions'].extend([
                    'Conduct post-hoc tests to identify specific group differences',
                    'Consider effect size calculations for practical significance',
                    'Validate findings with additional data if possible'
                ])
            else:
                conclusions['primary_conclusion'] = 'No significant differences were detected between the analyzed groups'
                conclusions['statistical_conclusion'] = f'The analysis does not provide sufficient evidence (p = {p_value:.4f}) to reject the null hypothesis'
                conclusions['future_directions'].extend([
                    'Consider increasing sample size for adequate power',
                    'Review measurement precision and data collection methods',
                    'Explore alternative analytical approaches if appropriate'
                ])
        
        return conclusions
    
    def _create_appendices(self, result: Dict[str, Any], validation_result: Dict[str, Any]) -> Dict[str, Any]:
        """Create appendices section"""
        appendices = {
            'title': 'Appendices',
            'sections': []
        }
        
        # Appendix A: Technical Details
        tech_details = {
            'title': 'Appendix A: Technical Details',
            'software_version': 'ANOVA Analysis Tool v2.0',
            'analysis_timestamp': self.export_timestamp.isoformat(),
            'calculation_precision': '15 decimal places',
            'statistical_packages': 'SciPy, NumPy, Pandas'
        }
        appendices['sections'].append(tech_details)
        
        # Appendix B: Validation Report
        if validation_result:
            validation_summary = {
                'title': 'Appendix B: Data Validation Report',
                'validation_status': validation_result.get('is_valid', False),
                'errors_found': len(validation_result.get('errors', [])),
                'warnings_issued': len(validation_result.get('warnings', [])),
                'export_ready': validation_result.get('export_ready', False)
            }
            appendices['sections'].append(validation_summary)
        
        return appendices
    
    def _create_professional_footer(self, validation_result: Dict[str, Any]) -> Dict[str, str]:
        """Create professional footer information"""
        return {
            'generated_by': 'ANOVA Analysis Tool v2.0',
            'timestamp': self.export_timestamp.strftime('%Y-%m-%d %H:%M:%S'),
            'validation_status': 'Validated' if validation_result.get('is_valid', False) else 'Validation Issues',
            'page_info': 'Professional Statistical Analysis Report',
            'disclaimer': 'This report was generated automatically. Please verify results and interpretations.'
        }
    
    # DataFrame creation methods for Excel export
    def _create_summary_dataframe(self, result: Dict[str, Any], validation_result: Dict[str, Any]) -> pd.DataFrame:
        """Create summary DataFrame for Excel export"""
        summary_data = []
        
        if 'anova' in result:
            anova = result['anova']
            summary_data.extend([
                ['Analysis Type', 'One-Way ANOVA'],
                ['F-Statistic', anova.get('fStatistic', 'N/A')],
                ['P-Value', anova.get('pValue', 'N/A')],
                ['Degrees of Freedom (Between)', anova.get('dfBetween', 'N/A')],
                ['Degrees of Freedom (Within)', anova.get('dfWithin', 'N/A')],
                ['Sum of Squares (Between)', anova.get('ssBetween', 'N/A')],
                ['Sum of Squares (Within)', anova.get('ssWithin', 'N/A')],
                ['Statistical Significance', 'Yes' if anova.get('pValue', 1) < 0.05 else 'No']
            ])
        
        summary_data.extend([
            ['Export Timestamp', self.export_timestamp.strftime('%Y-%m-%d %H:%M:%S')],
            ['Validation Status', 'Valid' if validation_result.get('is_valid', False) else 'Issues Found'],
            ['Format Version', self.format_version]
        ])
        
        return pd.DataFrame(summary_data, columns=['Statistic', 'Value'])
    
    def _create_anova_dataframe(self, anova: Dict[str, Any]) -> pd.DataFrame:
        """Create ANOVA table DataFrame"""
        if not anova:
            return pd.DataFrame({'Note': ['No ANOVA results available']})
        
        anova_data = {
            'Source': ['Between Groups', 'Within Groups', 'Total'],
            'DF': [
                anova.get('dfBetween', 'N/A'),
                anova.get('dfWithin', 'N/A'),
                anova.get('dfTotal', 'N/A')
            ],
            'Sum of Squares': [
                anova.get('ssBetween', 0),
                anova.get('ssWithin', 0),
                anova.get('ssTotal', 0)
            ],
            'Mean Square': [
                anova.get('msBetween', 0),
                anova.get('msWithin', 0),
                'N/A'
            ],
            'F-Ratio': [
                anova.get('fStatistic', 0),
                'N/A',
                'N/A'
            ],
            'P-Value': [
                anova.get('pValue', 0),
                'N/A',
                'N/A'
            ]
        }
        
        return pd.DataFrame(anova_data)
    
    def _create_descriptive_dataframe(self, means: Dict[str, Any]) -> pd.DataFrame:
        """Create descriptive statistics DataFrame"""
        if not means or 'groupMeans' not in means:
            return pd.DataFrame({'Note': ['No descriptive statistics available']})
        
        group_means = means['groupMeans']
        descriptive_data = []
        
        for group in sorted(group_means.keys()):
            row_data = {
                'Group': group,
                'N': means.get('groupCounts', {}).get(group, 'N/A'),
                'Mean': group_means[group],
                'Std Dev': means.get('groupStds', {}).get(group, 'N/A'),
                'Std Error': means.get('groupSE', {}).get(group, 'N/A'),
                '95% CI Lower': means.get('groupCILower', {}).get(group, 'N/A'),
                '95% CI Upper': means.get('groupCIUpper', {}).get(group, 'N/A')
            }
            descriptive_data.append(row_data)
        
        return pd.DataFrame(descriptive_data)
    
    def _create_tukey_dataframe(self, tukey: Dict[str, Any]) -> pd.DataFrame:
        """Create Tukey HSD DataFrame"""
        if 'pairwiseComparisons' not in tukey:
            return pd.DataFrame({'Note': ['No Tukey HSD results available']})
        
        comparisons = tukey['pairwiseComparisons']
        tukey_data = []
        
        for pair, data in comparisons.items():
            row_data = {
                'Comparison': pair,
                'Mean Difference': data.get('meanDiff', 0),
                'P-Value': data.get('pValue', 0),
                'Significant': 'Yes' if data.get('significant', False) else 'No',
                'CI Lower': data.get('ciLower', 0),
                'CI Upper': data.get('ciUpper', 0)
            }
            tukey_data.append(row_data)
        
        return pd.DataFrame(tukey_data)
    
    def _create_assumptions_dataframe(self, result: Dict[str, Any]) -> pd.DataFrame:
        """Create assumptions tests DataFrame"""
        assumptions_data = []
        
        if 'levene' in result:
            levene = result['levene']
            assumptions_data.append({
                'Test': "Levene's Test",
                'Purpose': 'Homogeneity of Variance',
                'Statistic': levene.get('statistic', 'N/A'),
                'P-Value': levene.get('pValue', 'N/A'),
                'Assumption Met': 'Yes' if levene.get('pValue', 0) > 0.05 else 'No'
            })
        
        if 'shapiro' in result:
            shapiro = result['shapiro']
            assumptions_data.append({
                'Test': 'Shapiro-Wilk Test',
                'Purpose': 'Normality',
                'Statistic': shapiro.get('statistic', 'N/A'),
                'P-Value': shapiro.get('pValue', 'N/A'),
                'Assumption Met': 'Yes' if shapiro.get('pValue', 0) > 0.05 else 'No'
            })
        
        if not assumptions_data:
            assumptions_data.append({
                'Test': 'No Tests Performed',
                'Purpose': 'N/A',
                'Statistic': 'N/A',
                'P-Value': 'N/A',
                'Assumption Met': 'Unknown'
            })
        
        return pd.DataFrame(assumptions_data)
    
    def _create_raw_data_dataframe(self, raw_groups: Dict[str, List]) -> pd.DataFrame:
        """Create raw data DataFrame from groups"""
        all_data = []
        
        for group_name, values in raw_groups.items():
            for value in values:
                all_data.append({
                    'Group': group_name,
                    'Value': value
                })
        
        return pd.DataFrame(all_data)
    
    def _create_validation_dataframe(self, validation_result: Dict[str, Any]) -> pd.DataFrame:
        """Create validation report DataFrame"""
        validation_data = [
            ['Validation Status', 'Valid' if validation_result.get('is_valid', False) else 'Invalid'],
            ['Export Ready', 'Yes' if validation_result.get('export_ready', False) else 'No'],
            ['Errors Found', len(validation_result.get('errors', []))],
            ['Warnings Issued', len(validation_result.get('warnings', []))],
            ['Validation Timestamp', validation_result.get('metadata', {}).get('validation_timestamp', 'N/A')]
        ]
        
        # Add errors and warnings details
        for i, error in enumerate(validation_result.get('errors', []), 1):
            validation_data.append([f'Error {i}', error])
        
        for i, warning in enumerate(validation_result.get('warnings', []), 1):
            validation_data.append([f'Warning {i}', warning])
        
        return pd.DataFrame(validation_data, columns=['Item', 'Details'])
    
    # CSV creation methods
    def _create_summary_csv(self, result: Dict[str, Any]) -> str:
        """Create summary CSV string"""
        csv_buffer = io.StringIO()
        csv_buffer.write("ANOVA Analysis Summary\n")
        csv_buffer.write(f"Generated: {self.export_timestamp.strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        
        if 'anova' in result:
            anova = result['anova']
            csv_buffer.write("Statistic,Value\n")
            csv_buffer.write(f"F-Statistic,{anova.get('fStatistic', 'N/A')}\n")
            csv_buffer.write(f"P-Value,{anova.get('pValue', 'N/A')}\n")
            csv_buffer.write(f"DF Between,{anova.get('dfBetween', 'N/A')}\n")
            csv_buffer.write(f"DF Within,{anova.get('dfWithin', 'N/A')}\n")
            csv_buffer.write(f"SS Between,{anova.get('ssBetween', 'N/A')}\n")
            csv_buffer.write(f"SS Within,{anova.get('ssWithin', 'N/A')}\n")
        
        return csv_buffer.getvalue()
    
    def _create_anova_csv(self, anova: Dict[str, Any]) -> str:
        """Create ANOVA table CSV string"""
        csv_buffer = io.StringIO()
        csv_buffer.write("ANOVA Table\n")
        csv_buffer.write("Source,DF,Sum of Squares,Mean Square,F-Ratio,P-Value\n")
        
        if anova:
            csv_buffer.write(f"Between Groups,{anova.get('dfBetween', 'N/A')},{anova.get('ssBetween', 0)},{anova.get('msBetween', 0)},{anova.get('fStatistic', 0)},{anova.get('pValue', 0)}\n")
            csv_buffer.write(f"Within Groups,{anova.get('dfWithin', 'N/A')},{anova.get('ssWithin', 0)},{anova.get('msWithin', 0)},N/A,N/A\n")
            csv_buffer.write(f"Total,{anova.get('dfTotal', 'N/A')},{anova.get('ssTotal', 0)},N/A,N/A,N/A\n")
        
        return csv_buffer.getvalue()
    
    def _create_means_csv(self, means: Dict[str, Any]) -> str:
        """Create means CSV string"""
        csv_buffer = io.StringIO()
        csv_buffer.write("Descriptive Statistics by Group\n")
        csv_buffer.write("Group,N,Mean,Std Dev,Std Error,95% CI Lower,95% CI Upper\n")
        
        if 'groupMeans' in means:
            group_means = means['groupMeans']
            for group in sorted(group_means.keys()):
                csv_buffer.write(f"{group},{means.get('groupCounts', {}).get(group, 'N/A')},{group_means[group]},{means.get('groupStds', {}).get(group, 'N/A')},{means.get('groupSE', {}).get(group, 'N/A')},{means.get('groupCILower', {}).get(group, 'N/A')},{means.get('groupCIUpper', {}).get(group, 'N/A')}\n")
        
        return csv_buffer.getvalue()
    
    def _create_tukey_csv(self, tukey: Dict[str, Any]) -> str:
        """Create Tukey HSD CSV string"""
        csv_buffer = io.StringIO()
        csv_buffer.write("Tukey HSD Post-Hoc Test Results\n")
        csv_buffer.write("Comparison,Mean Difference,P-Value,Significant,CI Lower,CI Upper\n")
        
        if 'pairwiseComparisons' in tukey:
            comparisons = tukey['pairwiseComparisons']
            for pair, data in comparisons.items():
                significant = 'Yes' if data.get('significant', False) else 'No'
                csv_buffer.write(f"{pair},{data.get('meanDiff', 0)},{data.get('pValue', 0)},{significant},{data.get('ciLower', 0)},{data.get('ciUpper', 0)}\n")
        
        return csv_buffer.getvalue()