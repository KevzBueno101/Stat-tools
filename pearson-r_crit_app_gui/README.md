# Pearson's r Critical Value Calculator - Excel Analysis Feature

## Overview

The application now includes **automated Excel file analysis** functionality! You can now upload an Excel file and automatically compute Pearson correlation coefficients along with critical values for significance testing.

## What's New

### New Files Created

1. **excel_utils.py** - Core Excel processing module with functions for:
   - Reading Excel files
   - Extracting numeric columns
   - Computing Pearson correlations from data
   - Validating data for correlation analysis
   - Complete end-to-end analysis pipeline

2. **Updated Files**:
   - **gui_components.py** - Added `ExcelAnalysisFrame` component
   - **main.py** - Added tabbed interface with Excel analysis tab

## Features

### Manual Calculation Tab (Original)
- Compute critical r-values given sample size and alpha
- Supports both one-tailed and two-tailed tests
- Save results to Word document

### Excel Analysis Tab (NEW)
- **Browse and load Excel files** (.xlsx, .xls)
- **Automatic detection** of numeric columns
- **Select two variables** from dropdown menus
- **Automated computation** of:
  - Pearson's r correlation coefficient
  - P-value
  - Critical r-value
  - Statistical significance determination
- **Comprehensive results** including:
  - Correlation strength interpretation (weak, moderate, strong)
  - Significance testing
  - Handles missing data automatically
- **Save to Word document** with full analysis report

## How to Use

### Excel File Requirements

Your Excel file should:
- Contain **at least 2 numeric columns** (numbers)
- Have **at least 3 rows of data**
- Column headers in the first row (recommended)

### Step-by-Step Instructions

1. **Launch the application**
   ```bash
   python main.py
   ```

2. **Navigate to "Excel Analysis" tab**

3. **Click "Browse Excel File"**
   - Select your Excel file (.xlsx or .xls)
   - The app will automatically detect numeric columns

4. **Select your variables**
   - Choose Column 1 (X variable) from the dropdown
   - Choose Column 2 (Y variable) from the dropdown
   - They must be different columns

5. **Set analysis parameters** (optional)
   - Significance Level (α): Default is 0.05
   - Test Type: Two-tailed (default) or One-tailed

6. **Click "Analyze File"**
   - The app will compute the correlation
   - Results appear in the right panel

7. **Review results** which include:
   - Pearson's r value
   - P-value
   - Critical r-value
   - Whether correlation is significant
   - Correlation strength interpretation

8. **Save results** (optional)
   - Click "Save Results" to export to Word document

## Example Excel File Structure

```
| Student_ID | Study_Hours | Test_Score | Age |
|------------|-------------|------------|-----|
| 1          | 5           | 85         | 20  |
| 2          | 3           | 72         | 19  |
| 3          | 8           | 95         | 21  |
| 4          | 6           | 88         | 20  |
| 5          | 4           | 78         | 19  |
```

You could analyze:
- Study_Hours vs Test_Score
- Age vs Test_Score
- Study_Hours vs Age

## Required Dependencies

Make sure you have these packages installed:

```bash
pip install customtkinter
pip install scipy
pip install numpy
pip install pandas
pip install openpyxl  # For .xlsx files
pip install xlrd      # For .xls files (optional)
pip install python-docx
```

## Technical Details

### Data Validation

The app automatically:
- Checks if columns exist and are numeric
- Removes rows with missing data (NaN values)
- Reports how many rows were excluded
- Validates minimum sample size (n ≥ 3)
- Checks for zero variance in variables

### Statistical Computations

1. **Pearson Correlation**: Uses `scipy.stats.pearsonr()`
2. **Critical Value**: Computed using t-distribution
3. **Significance Test**: Compares |r| to r_critical
4. **Strength Interpretation**:
   - |r| ≥ 0.7: Strong
   - |r| ≥ 0.4: Moderate
   - |r| ≥ 0.2: Weak
   - |r| < 0.2: Very weak

### Formula Used

```
r_critical = sqrt(t² / (t² + df))

where:
- t = critical t-value from t-distribution
- df = n - 2 (degrees of freedom)
- n = sample size
```

## Troubleshooting

### "File must contain at least 2 numeric columns"
- Ensure your Excel file has columns with numbers (not text)
- Check that column data types are numeric

### "Not enough valid data points"
- Need at least 3 complete pairs of data
- Check for too many missing values (empty cells)

### "Column has zero variance"
- All values in a column are identical
- Cannot compute correlation (mathematically undefined)

### Analysis button disabled
- Make sure you've selected a file first
- Verify the file loaded successfully
- Check that numeric columns were detected

## Tips for Best Results

1. **Clean your data** before importing:
   - Remove text from numeric columns
   - Handle missing values appropriately
   - Ensure consistent data types

2. **Use meaningful column names**:
   - Headers help identify variables
   - Makes results more interpretable

3. **Check assumptions**:
   - Linear relationship expected
   - Continuous variables
   - No extreme outliers

4. **Save your results**:
   - Word documents include full analysis
   - Great for reports and documentation

## Future Enhancements

Potential features for future versions:
- Scatter plot visualization
- Multiple correlation matrix
- Export to CSV/Excel
- Outlier detection
- Assumption checking

## Support

For issues or questions:
- Check that all dependencies are installed
- Verify Excel file format and structure
- Ensure Python 3.7+ is being used

---

**Enjoy the automated Excel analysis feature!** 📊✨