"""
ANOVA Results Validation Module
Provides comprehensive validation for ANOVA analysis results before export
เพื่อให้มั่นใจว่าข้อมูลที่ Export มีความถูกต้องและครบถ้วน
"""

import numpy as np
import pandas as pd
from typing import Dict, List, Tuple, Optional, Any
from datetime import datetime
import warnings


class ANOVAValidator:
    """
    Class สำหรับตรวจสอบความถูกต้องของผลลัพธ์ ANOVA
    """
    
    def __init__(self, alpha: float = 0.05):
        """
        Initialize validator with significance level
        
        Args:
            alpha: significance level (default 0.05)
        """
        self.alpha = alpha
        self.validation_errors = []
        self.validation_warnings = []
        
    def reset_validation_status(self):
        """Reset validation errors and warnings"""
        self.validation_errors.clear()
        self.validation_warnings.clear()
    
    def validate_anova_results(self, result: Dict[str, Any]) -> Dict[str, Any]:
        """
        Comprehensive validation of ANOVA results
        
        Args:
            result: Dictionary containing ANOVA analysis results
            
        Returns:
            Dictionary with validation status and enhanced metadata
        """
        self.reset_validation_status()
        
        validation_result = {
            'is_valid': True,
            'errors': [],
            'warnings': [],
            'metadata': {},
            'statistical_checks': {},
            'export_ready': False
        }
        
        # 1. Basic structure validation
        self._validate_basic_structure(result)
        
        # 2. ANOVA table validation
        if 'anova' in result:
            self._validate_anova_table(result['anova'])
        
        # 3. Means and group statistics validation
        if 'means' in result:
            self._validate_means_data(result['means'])
        
        # 4. Post-hoc tests validation
        if 'tukey' in result:
            self._validate_tukey_results(result['tukey'])
        
        # 5. Statistical assumptions checks
        self._check_statistical_assumptions(result)
        
        # 6. Generate metadata
        validation_result['metadata'] = self._generate_metadata(result)
        
        # Compile results
        validation_result['errors'] = self.validation_errors.copy()
        validation_result['warnings'] = self.validation_warnings.copy()
        validation_result['is_valid'] = len(self.validation_errors) == 0
        validation_result['export_ready'] = validation_result['is_valid'] and self._check_export_readiness(result)
        
        return validation_result
    
    def _validate_basic_structure(self, result: Dict[str, Any]):
        """Validate basic structure of results"""
        required_keys = ['anova', 'means', 'basicInfo']
        
        for key in required_keys:
            if key not in result:
                self.validation_errors.append(f"Missing required section: {key}")
        
        # Check if we have actual data
        if not result:
            self.validation_errors.append("Empty analysis results")
    
    def _validate_anova_table(self, anova: Dict[str, Any]):
        """Validate ANOVA table values"""
        required_anova_keys = ['fStatistic', 'pValue', 'dfBetween', 'dfWithin', 'ssBetween', 'ssWithin']
        
        for key in required_anova_keys:
            if key not in anova:
                self.validation_errors.append(f"Missing ANOVA statistic: {key}")
                continue
            
            value = anova[key]
            
            # Validate specific statistics
            if key == 'fStatistic':
                if not isinstance(value, (int, float)) or value < 0:
                    self.validation_errors.append(f"Invalid F-statistic: {value}")
                elif value == 0:
                    self.validation_warnings.append("F-statistic is zero - check data variability")
            
            elif key == 'pValue':
                if not isinstance(value, (int, float)) or not (0 <= value <= 1):
                    self.validation_errors.append(f"Invalid p-value: {value}")
                elif value < self.alpha:
                    # This is actually good news, but we log it
                    pass
            
            elif key in ['dfBetween', 'dfWithin']:
                if not isinstance(value, (int, float)) or value <= 0:
                    self.validation_errors.append(f"Invalid degrees of freedom ({key}): {value}")
            
            elif key in ['ssBetween', 'ssWithin']:
                if not isinstance(value, (int, float)) or value < 0:
                    self.validation_errors.append(f"Invalid sum of squares ({key}): {value}")
        
        # Cross-validate statistics
        if all(k in anova for k in ['ssBetween', 'ssWithin', 'dfBetween', 'dfWithin']):
            self._cross_validate_anova_statistics(anova)
    
    def _cross_validate_anova_statistics(self, anova: Dict[str, Any]):
        """Cross-validate ANOVA statistics for consistency"""
        try:
            ss_between = anova['ssBetween']
            ss_within = anova['ssWithin']
            df_between = anova['dfBetween']
            df_within = anova['dfWithin']
            
            # Calculate expected mean squares
            ms_between_expected = ss_between / df_between if df_between > 0 else 0
            ms_within_expected = ss_within / df_within if df_within > 0 else 0
            
            # Check if F-statistic matches
            f_expected = ms_between_expected / ms_within_expected if ms_within_expected > 0 else 0
            f_actual = anova.get('fStatistic', 0)
            
            if abs(f_expected - f_actual) > 0.001:  # Allow small rounding differences
                self.validation_warnings.append(f"F-statistic inconsistency: expected {f_expected:.6f}, got {f_actual:.6f}")
            
        except (ZeroDivisionError, KeyError) as e:
            self.validation_warnings.append(f"Could not cross-validate ANOVA statistics: {str(e)}")
    
    def _validate_means_data(self, means: Dict[str, Any]):
        """Validate group means and related statistics"""
        if 'groupMeans' not in means:
            self.validation_errors.append("Missing group means data")
            return
        
        group_means = means['groupMeans']
        if not isinstance(group_means, dict) or len(group_means) < 2:
            self.validation_errors.append("Insufficient group means data (need at least 2 groups)")
        
        # Validate individual means
        for group, mean in group_means.items():
            if not isinstance(mean, (int, float)):
                self.validation_errors.append(f"Invalid mean for group {group}: {mean}")
    
    def _validate_tukey_results(self, tukey: Dict[str, Any]):
        """Validate Tukey post-hoc test results"""
        if 'pairwiseComparisons' not in tukey:
            self.validation_warnings.append("Missing pairwise comparisons in Tukey results")
            return
        
        comparisons = tukey['pairwiseComparisons']
        if not isinstance(comparisons, dict):
            self.validation_errors.append("Invalid Tukey pairwise comparisons format")
            return
        
        # Check each comparison
        for pair, data in comparisons.items():
            if not isinstance(data, dict):
                continue
            
            required_keys = ['pValue', 'meanDiff', 'significant']
            for key in required_keys:
                if key not in data:
                    self.validation_warnings.append(f"Missing {key} in Tukey comparison {pair}")
    
    def _check_statistical_assumptions(self, result: Dict[str, Any]):
        """Check statistical assumptions for ANOVA"""
        assumptions = {
            'normality': 'unknown',
            'homogeneity': 'unknown',
            'independence': 'assumed'
        }
        
        # Check if we have Levene's test results
        if 'levene' in result:
            levene = result['levene']
            if 'pValue' in levene:
                p_val = levene['pValue']
                assumptions['homogeneity'] = 'satisfied' if p_val > self.alpha else 'violated'
                if p_val <= self.alpha:
                    self.validation_warnings.append(f"Levene's test suggests heterogeneity (p={p_val:.4f})")
        
        # Check if we have normality tests
        if 'shapiro' in result:
            shapiro = result['shapiro']
            if 'pValue' in shapiro:
                p_val = shapiro['pValue']
                assumptions['normality'] = 'satisfied' if p_val > self.alpha else 'violated'
                if p_val <= self.alpha:
                    self.validation_warnings.append(f"Shapiro-Wilk test suggests non-normality (p={p_val:.4f})")
        
        return assumptions
    
    def _generate_metadata(self, result: Dict[str, Any]) -> Dict[str, Any]:
        """Generate comprehensive metadata for export"""
        metadata = {
            'validation_timestamp': datetime.now().isoformat(),
            'validator_version': '1.0.0',
            'significance_level': self.alpha,
            'export_format_version': '2.0'
        }
        
        # Extract basic information
        if 'basicInfo' in result:
            basic_info = result['basicInfo']
            metadata.update({
                'total_observations': basic_info.get('totalN', 'Unknown'),
                'number_of_groups': basic_info.get('groupCount', 'Unknown'),
                'group_names': basic_info.get('lotNames', [])
            })
        
        # Extract statistical summary
        if 'anova' in result:
            anova = result['anova']
            metadata.update({
                'f_statistic': anova.get('fStatistic', 'N/A'),
                'p_value': anova.get('pValue', 'N/A'),
                'effect_significant': anova.get('pValue', 1) < self.alpha,
                'degrees_of_freedom': {
                    'between': anova.get('dfBetween', 'N/A'),
                    'within': anova.get('dfWithin', 'N/A'),
                    'total': anova.get('dfTotal', 'N/A')
                }
            })
        
        return metadata
    
    def _check_export_readiness(self, result: Dict[str, Any]) -> bool:
        """Check if results are ready for professional export"""
        if self.validation_errors:
            return False
        
        # Must have core ANOVA results
        if 'anova' not in result:
            return False
        
        anova = result['anova']
        required_stats = ['fStatistic', 'pValue', 'dfBetween', 'dfWithin']
        
        return all(stat in anova for stat in required_stats)


def validate_export_data(result: Dict[str, Any], alpha: float = 0.05) -> Dict[str, Any]:
    """
    Convenience function for validating ANOVA export data
    
    Args:
        result: ANOVA analysis results
        alpha: significance level
        
    Returns:
        Validation results with metadata
    """
    validator = ANOVAValidator(alpha=alpha)
    return validator.validate_anova_results(result)


def create_export_summary(result: Dict[str, Any], validation_result: Dict[str, Any]) -> Dict[str, Any]:
    """
    Create executive summary for export
    
    Args:
        result: ANOVA analysis results
        validation_result: Validation results from validate_export_data
        
    Returns:
        Dictionary containing executive summary
    """
    summary = {
        'title': 'ANOVA Analysis Executive Summary',
        'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'status': 'Valid' if validation_result['is_valid'] else 'Issues Found',
        'overview': {},
        'key_findings': [],
        'recommendations': []
    }
    
    # Extract key statistics
    if 'anova' in result and validation_result['is_valid']:
        anova = result['anova']
        p_value = anova.get('pValue', 1)
        f_stat = anova.get('fStatistic', 0)
        
        summary['overview'] = {
            'test_type': 'One-Way ANOVA',
            'significance_level': validation_result.get('metadata', {}).get('significance_level', 0.05),
            'f_statistic': f_stat,
            'p_value': p_value,
            'result': 'Significant' if p_value < 0.05 else 'Not Significant'
        }
        
        # Key findings
        if p_value < 0.05:
            summary['key_findings'].append(f"Significant difference found between groups (F={f_stat:.4f}, p={p_value:.4f})")
            summary['recommendations'].append("Proceed with post-hoc tests to identify which groups differ")
        else:
            summary['key_findings'].append(f"No significant difference between groups (F={f_stat:.4f}, p={p_value:.4f})")
            summary['recommendations'].append("Groups do not differ significantly at the chosen significance level")
    
    # Add validation warnings if any
    if validation_result['warnings']:
        summary['key_findings'].extend([f"Warning: {w}" for w in validation_result['warnings']])
    
    return summary