"""
Statistical Utilities Module
Contains reusable functions for statistical computations related to Pearson's r.
"""

from scipy.stats import t
import numpy as np
from typing import Dict, Tuple


def validate_sample_size(n: float) -> Tuple[bool, str]:
    """
    Validate sample size input.
    
    Parameters:
    -----------
    n : float
        Sample size to validate
        
    Returns:
    --------
    tuple : (bool, str)
        (is_valid, error_message)
    """
    try:
        n = float(n)
        if n < 3:
            return False, "Sample size must be at least 3"
        if not n.is_integer():
            return False, "Sample size must be a whole number"
        return True, ""
    except (ValueError, TypeError):
        return False, "Sample size must be a valid number"


def validate_alpha(alpha: float) -> Tuple[bool, str]:
    """
    Validate alpha (significance level) input.
    
    Parameters:
    -----------
    alpha : float
        Significance level to validate
        
    Returns:
    --------
    tuple : (bool, str)
        (is_valid, error_message)
    """
    try:
        alpha = float(alpha)
        if alpha <= 0 or alpha >= 1:
            return False, "Alpha must be between 0 and 1 (exclusive)"
        return True, ""
    except (ValueError, TypeError):
        return False, "Alpha must be a valid number"


def compute_degrees_of_freedom(n: int) -> int:
    """
    Compute degrees of freedom for Pearson correlation.
    
    For Pearson's r, df = n - 2 because we estimate two parameters
    (slope and intercept in the underlying regression model).
    
    Parameters:
    -----------
    n : int
        Sample size
        
    Returns:
    --------
    int : Degrees of freedom
    """
    return n - 2


def compute_t_critical(alpha: float, df: int, test_type: str = "two-tailed") -> float:
    """
    Compute the critical t-value for given alpha and degrees of freedom.
    
    Uses the t-distribution's percent point function (inverse CDF).
    For two-tailed test, we use alpha/2 in each tail.
    For one-tailed test, we use alpha in one tail.
    
    Parameters:
    -----------
    alpha : float
        Significance level (e.g., 0.05)
    df : int
        Degrees of freedom
    test_type : str
        Either "two-tailed" or "one-tailed"
        
    Returns:
    --------
    float : Critical t-value
    """
    if test_type == "two-tailed":
        # For two-tailed, split alpha between both tails
        t_crit = t.ppf(1 - alpha / 2, df)
    else:
        # For one-tailed, all alpha in one tail
        t_crit = t.ppf(1 - alpha, df)
    
    return t_crit


def compute_pearson_r_critical(n: int, alpha: float = 0.05, 
                               test_type: str = "two-tailed") -> Dict:
    """
    Compute the critical value of Pearson's correlation coefficient.
    
    The relationship between t-statistic and Pearson's r is:
    t = r * sqrt(n - 2) / sqrt(1 - r²)
    
    Solving for r in terms of t:
    r = sqrt(t² / (t² + df))
    
    where df = n - 2
    
    Parameters:
    -----------
    n : int
        Sample size
    alpha : float
        Significance level (default: 0.05)
    test_type : str
        Either "two-tailed" or "one-tailed" (default: "two-tailed")
        
    Returns:
    --------
    dict : Dictionary containing:
        - sample_size: Input sample size
        - alpha: Significance level
        - test_type: Type of test
        - degrees_of_freedom: Computed df
        - t_critical: Critical t-value
        - r_critical: Critical r-value
    """
    # Compute degrees of freedom
    df = compute_degrees_of_freedom(n)
    
    # Compute critical t-value
    t_crit = compute_t_critical(alpha, df, test_type)
    
    # Convert t-critical to r-critical using the formula:
    # r = sqrt(t² / (t² + df))
    # This formula comes from the relationship between the t-statistic
    # and Pearson's r in hypothesis testing
    r_crit = np.sqrt(t_crit**2 / (t_crit**2 + df))
    
    # Return all results as a dictionary for reusability
    return {
        'sample_size': n,
        'alpha': alpha,
        'test_type': test_type,
        'degrees_of_freedom': df,
        't_critical': t_crit,
        'r_critical': r_crit
    }


def format_results(results: Dict, decimal_places: int = 6) -> str:
    """
    Format results dictionary into a readable string.
    
    Parameters:
    -----------
    results : dict
        Results dictionary from compute_pearson_r_critical
    decimal_places : int
        Number of decimal places for formatting (default: 6)
        
    Returns:
    --------
    str : Formatted results string
    """
    output = []
    output.append("=" * 50)
    output.append("PEARSON'S R CRITICAL VALUE RESULTS")
    output.append("=" * 50)
    output.append(f"Sample Size (n): {results['sample_size']}")
    output.append(f"Significance Level (α): {results['alpha']}")
    output.append(f"Test Type: {results['test_type'].title()}")
    output.append(f"Degrees of Freedom (df): {results['degrees_of_freedom']}")
    output.append(f"t Critical: {results['t_critical']:.{decimal_places}f}")
    output.append(f"r Critical: {results['r_critical']:.{decimal_places}f}")
    output.append("=" * 50)
    output.append(f"\nInterpretation: For a correlation to be statistically")
    output.append(f"significant at α = {results['alpha']}, the absolute value")
    output.append(f"of r must be greater than {results['r_critical']:.4f}")
    
    return "\n".join(output)