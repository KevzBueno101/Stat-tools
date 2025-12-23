from scipy.stats import t
import numpy as np

# Example data
r = 0.8          # Pearson r
n = 10           # sample size
alpha = 0.05     # significance level

df = n - 2

# t-statistic from r
t_stat = r * np.sqrt(df / (1 - r**2))

# critical t-value (two-tailed)
t_crit = t.ppf(1 - alpha/2, df)

print("t-statistic:", t_stat)
print("t-critical:", t_crit)

if abs(t_stat) > t_crit:
    print("Reject H₀ → significant correlation")
else:
    print("Fail to reject H₀ → not significant")
