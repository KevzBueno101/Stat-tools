# Correlation coefficient calculation using scipy
# in Excel = CORREL(array1, array2)

import numpy as np
from scipy.stats import pearsonr

x = [1, 2, 3, 4, 5]
y = [2, 3, 5, 7, 9]

r, p = pearsonr(x, y)
print(f"Computed value: {r}")  # observed r
