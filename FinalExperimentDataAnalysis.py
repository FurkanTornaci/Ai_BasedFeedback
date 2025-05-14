import pandas as pd
import numpy as np
from scipy.stats import friedmanchisquare, wilcoxon
import matplotlib.pyplot as plt
import seaborn as sns
import warnings
from sklearn.preprocessing import OrdinalEncoder

# Optional: suppress warning for Wilcoxon with zero differences
warnings.filterwarnings("ignore", message="Exact p-value calculation does not work if there are zeros")

# -------------------------------
# Load and prepare data
# -------------------------------
file_path = 'FinalExperimentDataWith60Particiapants.csv'
df_raw = pd.read_csv(file_path)

# Use row 0 as header, skip metadata rows
df = df_raw.iloc[2:].reset_index(drop=True)
df.columns = df_raw.iloc[0]
df = df.reset_index(drop=True)
print(f"After header fix and metadata skip: {df.shape[0]} rows")


# Keep only "Finished" responses
before = df.shape[0]
df = df[df["Finished"].astype(str).str.strip().str.lower() == "true"].copy()
print(f"Removed {before - df.shape[0]} unfinished responses; Remaining: {df.shape[0]}")

duration_col = "Duration (in seconds)"
df[duration_col] = pd.to_numeric(df[duration_col], errors='coerce')

# -------------------------------
# Remove low-duration responses (below 10th percentile)
# -------------------------------
duration_threshold = df[duration_col].quantile(0.1)
before = df.shape[0]
df = df[df[duration_col] >= duration_threshold].reset_index(drop=True)
print(f"Removed {before - df.shape[0]} low-duration responses (< 10th percentile); Remaining: {df.shape[0]}")

# -------------------------------
# Define the feedback dimensions
# -------------------------------
dimensions = ["Sufficiency", "Usefulness", "Clarity", "Adaptiveness", "Motivational Impact"]

# -------------------------------
# Function to get the column name based on feedback type and dimension
# -------------------------------
def get_column(feedback_type, dimension):
    for col in df.columns:
        if feedback_type in col and dimension in col:
            return col
    return None

# -------------------------------
# Statistical Tests (Friedman + Wilcoxon)
# -------------------------------
ordinal_mapping = {"Very Low": 1, "Low": 2, "Medium": 3, "High": 4, "Very High": 5}

friedman_results = []
wilcoxon_results = []

for dimension in dimensions:
    col_a = get_column("Feedback A", dimension)
    col_b = get_column("Feedback B", dimension)
    col_c = get_column("Feedback C", dimension)

    if col_a and col_b and col_c:
        data = df[[col_a, col_b, col_c]].apply(lambda col: col.map(ordinal_mapping.get)).dropna()
        data.columns = ["A", "B", "C"]

        # Friedman test
        f_stat, f_pval = friedmanchisquare(data["A"], data["B"], data["C"])
        friedman_results.append({
            "Metric": f"Friedman - {dimension}",
            "Statistic": round(f_stat, 3),
            "p-value": f"{round(f_pval, 4)}{' *' if f_pval < 0.05 else ''}",
            "Significant": f_pval < 0.05
        })

        # Wilcoxon pairwise tests
        for pair in [("A", "B"), ("A", "C"), ("B", "C")]:
            try:
                stat, p_val = wilcoxon(data[pair[0]], data[pair[1]])
                diffs = data[pair[0]] - data[pair[1]]
                median_diff = diffs.median()
                n_pos = (diffs > 0).sum()
                n_neg = (diffs < 0).sum()

                if median_diff > 0:
                    better = pair[0]
                elif median_diff < 0:
                    better = pair[1]
                elif n_pos > n_neg:
                    better = pair[0]
                elif n_neg > n_pos:
                    better = pair[1]
                else:
                    better = "Equal"

                wilcoxon_results.append({
                    "Metric": f"Wilcoxon {pair[0]} vs {pair[1]} - {dimension}",
                    "W-stat": round(stat, 3),
                    "p-value": f"{round(p_val, 4)}{' *' if p_val < 0.05 else ''}",
                    "Significant": p_val < 0.05,
                    "Better": better
                })
            except ValueError:
                wilcoxon_results.append({
                    "Metric": f"Wilcoxon {pair[0]} vs {pair[1]} - {dimension}",
                    "W-stat": "NA",
                    "p-value": "NA",
                    "Significant": False,
                    "Better": "NA"
                })

# -------------------------------
# Reliability Analysis (Cronbach's Alpha)
# -------------------------------
def cronbach_alpha(df_sub):
    # Drop rows with missing values
    df_sub = df_sub.dropna()
    item_scores = df_sub.values
    item_vars = item_scores.var(axis=0, ddof=1)
    total_var = item_scores.sum(axis=1).var(ddof=1)
    n_items = item_scores.shape[1]
    if total_var == 0:
        return np.nan
    return (n_items / (n_items - 1)) * (1 - item_vars.sum() / total_var)

# Prepare ordinal encoding for Likert scale
encoder = OrdinalEncoder(categories=[["Very Low", "Low", "Medium", "High", "Very High"]] * len(dimensions))

# Reliability Analysis for each feedback type
reliability_scores = []

for fb_type in ["Feedback A", "Feedback B", "Feedback C"]:
    cols = [get_column(fb_type, dim) for dim in dimensions]
    cols = [col for col in cols if col]
    df_sub = df[cols].dropna()  # Ensure no missing values
    df_encoded = pd.DataFrame(encoder.fit_transform(df_sub), columns=cols)
    alpha = cronbach_alpha(df_encoded)
    reliability_scores.append({
        "Feedback Type": fb_type,
        "Cronbach's Alpha": round(alpha, 3)
    })

# Display reliability analysis results
print("\n" + "="*50)
print("Reliability Analysis: Cronbach’s Alpha")
print("="*50)
for score in reliability_scores:
    alpha_val = score["Cronbach's Alpha"]
    print(f"{score['Feedback Type']:<12} | Alpha = {alpha_val}")
print("="*50)

# -------------------------------
# Display Test Results
# -------------------------------
n_participants = df.shape[0]

def format_table(results_df, title, stat_col_name):
    print("\n" + "=" * 120)
    print(f"{title} (n = {n_participants})".center(120))
    print("=" * 120)
    header = f"{'Metric':<50} | {stat_col_name:<8} | {'p-value':<10} | {'Better':<10} | Mean A  | Mean B  | Mean C"
    print(header)
    print("-" * len(header))
    for _, row in results_df.iterrows():
        line = f"{row['Metric']:<50} | {str(row.get(stat_col_name, '')):<8} | {str(row.get('p-value', '')).ljust(10)} | {str(row.get('Better', '')).ljust(10)}"
        line += f" | {str(row.get('Mean A', '')).ljust(7)} | {str(row.get('Mean B', '')).ljust(7)} | {str(row.get('Mean C', '')).ljust(7)}"
        print(line)
    print("=" * 120)

format_table(pd.DataFrame(friedman_results), "Friedman Test", "Statistic")
format_table(pd.DataFrame(wilcoxon_results), "Wilcoxon Signed-Rank Tests", "W-stat")

# -------------------------------
# Plot: Most Beneficial Feedback
# -------------------------------
beneficial_col = [col for col in df.columns if "Which Feedback Was Most Beneficial" in col][0]
plt.figure(figsize=(8, 5))
sns.countplot(data=df, y=beneficial_col, order=df[beneficial_col].value_counts().index, palette="Set3")
plt.title("Most Beneficial Feedback Type")
plt.xlabel("Number of Responses")
plt.ylabel("Feedback Type")
plt.tight_layout()
plt.show()
