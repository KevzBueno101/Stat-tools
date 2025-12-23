"""
Excel Utilities Module
Contains functions for reading and analyzing Excel files for correlation analysis.
"""

import pandas as pd
import numpy as np
from typing import Dict, List, Tuple, Optional
from scipy.stats import pearsonr


def read_excel_file(filepath: str) -> Tuple[bool, pd.DataFrame, str]:
    """
    Read an Excel file and return its contents.
    
    Parameters:
    -----------
    filepath : str
        Path to the Excel file
        
    Returns:
    --------
    tuple : (success, dataframe, error_message)
    """
    try:
        df = pd.read_excel(filepath)
        
        if df.empty:
            return False, None, "The Excel file is empty"
        
        return True, df, ""
    
    except FileNotFoundError:
        return False, None, "File not found"
    except Exception as e:
        return False, None, f"Error reading Excel file: {str(e)}"


def get_numeric_columns(df: pd.DataFrame) -> List[str]:
    """
    Get list of numeric columns from dataframe.
    
    Parameters:
    -----------
    df : pd.DataFrame
        Input dataframe
        
    Returns:
    --------
    list : List of numeric column names
    """
    return df.select_dtypes(include=[np.number]).columns.tolist()


def validate_columns_for_correlation(df: pd.DataFrame, col1: str, col2: str) -> Tuple[bool, str]:
    """
    Validate that two columns are suitable for correlation analysis.
    
    Parameters:
    -----------
    df : pd.DataFrame
        Input dataframe
    col1, col2 : str
        Column names to validate
        
    Returns:
    --------
    tuple : (is_valid, error_message)
    """
    # Check if columns exist
    if col1 not in df.columns:
        return False, f"Column '{col1}' not found in the Excel file"
    
    if col2 not in df.columns:
        return False, f"Column '{col2}' not found in the Excel file"
    
    # Check if columns are numeric
    if not np.issubdtype(df[col1].dtype, np.number):
        return False, f"Column '{col1}' is not numeric"
    
    if not np.issubdtype(df[col2].dtype, np.number):
        return False, f"Column '{col2}' is not numeric"
    
    # Drop NaN values for analysis
    clean_data = df[[col1, col2]].dropna()
    
    # Check if we have enough data points
    if len(clean_data) < 3:
        return False, f"Not enough valid data points (need at least 3, found {len(clean_data)})"
    
    # Check for zero variance
    if clean_data[col1].std() == 0:
        return False, f"Column '{col1}' has zero variance (all values are the same)"
    
    if clean_data[col2].std() == 0:
        return False, f"Column '{col2}' has zero variance (all values are the same)"
    
    return True, ""


def compute_correlation_from_data(df: pd.DataFrame, col1: str, col2: str) -> Dict:
    """
    Compute Pearson correlation coefficient from two columns of data.
    
    Parameters:
    -----------
    df : pd.DataFrame
        Input dataframe
    col1, col2 : str
        Column names to correlate
        
    Returns:
    --------
    dict : Dictionary containing correlation results
    """
    # Remove rows with NaN values in either column
    clean_data = df[[col1, col2]].dropna()
    
    # Extract the two variables
    x = clean_data[col1].values
    y = clean_data[col2].values
    
    # Compute Pearson correlation
    r_value, p_value = pearsonr(x, y)
    
    # Get sample size
    n = len(x)
    
    return {
        'column_1': col1,
        'column_2': col2,
        'sample_size': n,
        'r_value': r_value,
        'p_value': p_value,
        'original_rows': len(df),
        'rows_with_missing': len(df) - n
    }


def analyze_excel_correlation(filepath: str, col1: str, col2: str, 
                              alpha: float = 0.05, 
                              test_type: str = "two-tailed") -> Tuple[bool, Dict, str]:
    """
    Complete analysis pipeline: read Excel file, compute correlation, and determine significance.
    
    Parameters:
    -----------
    filepath : str
        Path to Excel file
    col1, col2 : str
        Column names to correlate
    alpha : float
        Significance level
    test_type : str
        Type of test ("two-tailed" or "one-tailed")
        
    Returns:
    --------
    tuple : (success, results_dict, error_message)
    """
    from stats_utils import compute_pearson_r_critical
    
    # Read Excel file
    success, df, error = read_excel_file(filepath)
    if not success:
        return False, None, error
    
    # Validate columns
    is_valid, error = validate_columns_for_correlation(df, col1, col2)
    if not is_valid:
        return False, None, error
    
    # Compute correlation from data
    corr_results = compute_correlation_from_data(df, col1, col2)
    
    # Compute critical value
    critical_results = compute_pearson_r_critical(
        corr_results['sample_size'], 
        alpha, 
        test_type
    )
    
    # Determine if correlation is significant
    is_significant = abs(corr_results['r_value']) > critical_results['r_critical']
    
    # Combine results
    final_results = {
        **corr_results,
        **critical_results,
        'is_significant': is_significant,
        'significance_interpretation': (
            f"The correlation is {'SIGNIFICANT' if is_significant else 'NOT SIGNIFICANT'} "
            f"at α = {alpha} (|r| = {abs(corr_results['r_value']):.4f} "
            f"{'>' if is_significant else '<'} {critical_results['r_critical']:.4f})"
        )
    }
    
    return True, final_results, ""


def get_all_correlation_pairs(df: pd.DataFrame, alpha: float = 0.05, 
                              test_type: str = "two-tailed") -> List[Dict]:
    """
    Compute correlations for all pairs of numeric columns.
    
    Parameters:
    -----------
    df : pd.DataFrame
        Input dataframe
    alpha : float
        Significance level
    test_type : str
        Type of test
        
    Returns:
    --------
    list : List of dictionaries containing correlation results for each pair
    """
    from stats_utils import compute_pearson_r_critical
    from itertools import combinations
    
    numeric_cols = get_numeric_columns(df)
    
    if len(numeric_cols) < 2:
        return []
    
    results = []
    
    # Generate all unique pairs
    for col1, col2 in combinations(numeric_cols, 2):
        try:
            # Validate and compute
            is_valid, _ = validate_columns_for_correlation(df, col1, col2)
            if not is_valid:
                continue
            
            corr_results = compute_correlation_from_data(df, col1, col2)
            critical_results = compute_pearson_r_critical(
                corr_results['sample_size'], 
                alpha, 
                test_type
            )
            
            is_significant = abs(corr_results['r_value']) > critical_results['r_critical']
            
            results.append({
                **corr_results,
                **critical_results,
                'is_significant': is_significant
            })
        
        except Exception:
            continue
    
    return results