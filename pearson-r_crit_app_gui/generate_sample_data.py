"""
Sample Data Generator
Creates sample Excel files for testing the Pearson correlation calculator.
"""

import pandas as pd
import numpy as np
from datetime import datetime


def generate_student_performance_data(n=50, output_file="sample_student_data.xlsx"):
    """
    Generate sample student performance data with correlated variables.
    
    Parameters:
    -----------
    n : int
        Number of students (default: 50)
    output_file : str
        Output filename
    """
    np.random.seed(42)  # For reproducibility
    
    # Generate study hours (2-10 hours)
    study_hours = np.random.uniform(2, 10, n)
    
    # Generate test scores correlated with study hours
    # Score = 50 + 4*study_hours + random noise
    test_scores = 50 + 4 * study_hours + np.random.normal(0, 5, n)
    test_scores = np.clip(test_scores, 0, 100)  # Clip to valid range
    
    # Generate attendance (correlated with test scores)
    attendance = 70 + 0.2 * test_scores + np.random.normal(0, 5, n)
    attendance = np.clip(attendance, 50, 100)
    
    # Generate age (mostly independent)
    age = np.random.randint(18, 25, n)
    
    # Generate sleep hours (weakly negatively correlated with study hours)
    sleep_hours = 8.5 - 0.15 * study_hours + np.random.normal(0, 1, n)
    sleep_hours = np.clip(sleep_hours, 4, 10)
    
    # Create DataFrame
    df = pd.DataFrame({
        'Student_ID': range(1, n + 1),
        'Study_Hours': np.round(study_hours, 1),
        'Test_Score': np.round(test_scores, 1),
        'Attendance_Percent': np.round(attendance, 1),
        'Age': age,
        'Sleep_Hours': np.round(sleep_hours, 1)
    })
    
    # Add some missing values randomly
    missing_indices = np.random.choice(n, size=int(n * 0.05), replace=False)
    df.loc[missing_indices, 'Test_Score'] = np.nan
    
    # Save to Excel
    df.to_excel(output_file, index=False)
    print(f"✓ Generated {output_file}")
    print(f"  - Rows: {n}")
    print(f"  - Columns: {len(df.columns)}")
    print(f"  - Expected correlations:")
    print(f"    • Study_Hours vs Test_Score: Strong positive")
    print(f"    • Test_Score vs Attendance_Percent: Moderate positive")
    print(f"    • Study_Hours vs Sleep_Hours: Weak negative")


def generate_sales_data(n=100, output_file="sample_sales_data.xlsx"):
    """
    Generate sample sales and marketing data.
    
    Parameters:
    -----------
    n : int
        Number of data points (default: 100)
    output_file : str
        Output filename
    """
    np.random.seed(123)
    
    # Generate marketing spend ($1000 - $10000)
    marketing_spend = np.random.uniform(1000, 10000, n)
    
    # Generate sales (correlated with marketing spend)
    sales = 5000 + 3.5 * marketing_spend + np.random.normal(0, 5000, n)
    sales = np.maximum(sales, 0)  # No negative sales
    
    # Generate website traffic (correlated with both)
    website_traffic = 500 + 0.5 * marketing_spend + np.random.normal(0, 1000, n)
    website_traffic = np.maximum(website_traffic, 100)
    
    # Generate customer satisfaction (weakly correlated)
    customer_satisfaction = 7 + 0.0001 * sales + np.random.normal(0, 1, n)
    customer_satisfaction = np.clip(customer_satisfaction, 1, 10)
    
    # Generate number of employees (mostly independent)
    employees = np.random.randint(5, 50, n)
    
    df = pd.DataFrame({
        'Month_ID': range(1, n + 1),
        'Marketing_Spend_USD': np.round(marketing_spend, 2),
        'Sales_USD': np.round(sales, 2),
        'Website_Traffic': np.round(website_traffic, 0).astype(int),
        'Customer_Satisfaction': np.round(customer_satisfaction, 1),
        'Employees': employees
    })
    
    df.to_excel(output_file, index=False)
    print(f"\n✓ Generated {output_file}")
    print(f"  - Rows: {n}")
    print(f"  - Columns: {len(df.columns)}")
    print(f"  - Expected correlations:")
    print(f"    • Marketing_Spend_USD vs Sales_USD: Strong positive")
    print(f"    • Marketing_Spend_USD vs Website_Traffic: Moderate positive")
    print(f"    • Sales_USD vs Customer_Satisfaction: Weak positive")


def generate_health_data(n=75, output_file="sample_health_data.xlsx"):
    """
    Generate sample health and fitness data.
    
    Parameters:
    -----------
    n : int
        Number of participants (default: 75)
    output_file : str
        Output filename
    """
    np.random.seed(456)
    
    # Generate exercise hours per week (0-15)
    exercise_hours = np.random.uniform(0, 15, n)
    
    # Generate BMI (negatively correlated with exercise)
    bmi = 30 - 0.8 * exercise_hours + np.random.normal(0, 2, n)
    bmi = np.clip(bmi, 18, 40)
    
    # Generate cholesterol (negatively correlated with exercise)
    cholesterol = 220 - 3 * exercise_hours + np.random.normal(0, 15, n)
    cholesterol = np.clip(cholesterol, 150, 300)
    
    # Generate blood pressure (weakly correlated)
    blood_pressure = 130 - 0.5 * exercise_hours + np.random.normal(0, 10, n)
    blood_pressure = np.clip(blood_pressure, 100, 180)
    
    # Generate age (independent)
    age = np.random.randint(25, 70, n)
    
    # Generate water intake (positively correlated with exercise)
    water_intake = 6 + 0.15 * exercise_hours + np.random.normal(0, 1, n)
    water_intake = np.clip(water_intake, 2, 12)
    
    df = pd.DataFrame({
        'Participant_ID': range(1, n + 1),
        'Exercise_Hours_Week': np.round(exercise_hours, 1),
        'BMI': np.round(bmi, 1),
        'Cholesterol_mg_dL': np.round(cholesterol, 1),
        'Blood_Pressure_mmHg': np.round(blood_pressure, 0).astype(int),
        'Age': age,
        'Water_Intake_Liters': np.round(water_intake, 1)
    })
    
    # Add some missing values
    missing_indices = np.random.choice(n, size=int(n * 0.08), replace=False)
    df.loc[missing_indices[:len(missing_indices)//2], 'Cholesterol_mg_dL'] = np.nan
    df.loc[missing_indices[len(missing_indices)//2:], 'BMI'] = np.nan
    
    df.to_excel(output_file, index=False)
    print(f"\n✓ Generated {output_file}")
    print(f"  - Rows: {n}")
    print(f"  - Columns: {len(df.columns)}")
    print(f"  - Expected correlations:")
    print(f"    • Exercise_Hours_Week vs BMI: Strong negative")
    print(f"    • Exercise_Hours_Week vs Cholesterol_mg_dL: Moderate negative")
    print(f"    • Exercise_Hours_Week vs Water_Intake_Liters: Moderate positive")


def generate_all_samples():
    """Generate all sample datasets."""
    print("=" * 60)
    print("Generating Sample Excel Files for Testing")
    print("=" * 60)
    
    generate_student_performance_data()
    generate_sales_data()
    generate_health_data()
    
    print("\n" + "=" * 60)
    print("✓ All sample files generated successfully!")
    print("=" * 60)
    print("\nYou can now use these files to test the Excel analysis feature:")
    print("  1. sample_student_data.xlsx")
    print("  2. sample_sales_data.xlsx")
    print("  3. sample_health_data.xlsx")
    print("\nOpen the Pearson Calculator app and try the 'Excel Analysis' tab!")


if __name__ == "__main__":
    generate_all_samples()