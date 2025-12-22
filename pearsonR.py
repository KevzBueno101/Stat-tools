#Computing the critical value of Pearson's r for a two-tailed test        
from scipy.stats import t
import numpy as np

alpha = 0.05
n = 99
# sample size
df = n - 2 #degrees of freedom

t_crit = t.ppf(1 - alpha/2, df) # critical value
r_crit = np.sqrt(t_crit**2 / (t_crit**2 + df))

print("Critical value of Pearson r:", r_crit)

#n = 340: Crit: 0.106
#n = 60: Crit: 0.254
#n = 69: Crit: 0.237
#n = 167: Crit: 0.152
#n = 90: Crit: 0.207
