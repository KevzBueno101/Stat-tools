import numpy as np
import pandas as pd
from statsmodels.stats.multicomp import pairwise_tukeyhsd

# Example: flatten groups into one array and create labels
group1 = [3.27, 3.47, 3.53, 3.27, 3.6]
group2 = [3, 3.67, 2.66, 2.66, 2.66]
group3 = [3.67, 3.8, 3.67 , 3.33, 3.67]

data = group1 + group2 + group3
labels = ['G1']*len(group1) + ['G2']*len(group2) + ['G3']*len(group3)

# Perform Tukey HSD
tukey = pairwise_tukeyhsd(endog=data, groups=labels, alpha=0.05)
print(tukey)

