def weighted_mean():
    print("Enter the frequencies for each rating:")
    
    f4 = int(input("Frequency for rating 4: "))
    f3 = int(input("Frequency for rating 3: "))
    f2 = int(input("Frequency for rating 2: "))
    f1 = int(input("Frequency for rating 1: "))

    # Weighted sum
    weighted_sum = (4 * f4) + (3 * f3) + (2 * f2) + (1 * f1)

    # Total respondents
    total_f = f4 + f3 + f2 + f1

    # Weighted mean
    if total_f == 0:
        print("No respondents. Weighted mean cannot be computed.")
    else:
        w_mean = weighted_sum / total_f
        print(f"\nWeighted Mean = {w_mean:.2f}")

# Run the function
weighted_mean()