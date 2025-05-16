
import pandas as pd
import numpy as np
from scipy.stats import friedmanchisquare, wilcoxon
from sklearn.preprocessing import OrdinalEncoder
import warnings

warnings.filterwarnings("ignore", message="Exact p-value calculation does not work if there are zeros")

# ----------------------------------
# Load Data
# ----------------------------------
file_path = 'Ai Feedback Assessment - Ready Survey_May 16, 2025_10.01.csv'
df_raw = pd.read_csv(file_path)

# Process headers
metadata_headers = df_raw.iloc[0]
question_labels = df_raw.iloc[1]
final_headers = [str(metadata_headers.iloc[i]).strip() if str(metadata_headers.iloc[i]).strip() else str(question_labels.iloc[i]).strip() for i in range(len(metadata_headers))]
df = df_raw.iloc[2:].reset_index(drop=True)
df.columns = final_headers

# ----------------------------------
# Filters
# ----------------------------------
df = df[df["Finished"].astype(str).str.strip().str.lower() == "true"].copy()

# Attention check
attention_check_col = None
for col in df.columns:
    if "attention" in str(col).lower() and "check" in str(col).lower():
        attention_check_col = col
        break
if attention_check_col is None:
    raise ValueError("Attention check column not found.")
df = df[df[attention_check_col].astype(str).str.strip().str.lower() == "low"].copy()

# Correct answers
correct_answers = {
    "Q1": 'with propelling machinery in use',
    "Q2": 'from the nature of her work is unable to maneuver as required by the rules',
    "Q3": 'A vessel engaged in fishing while underway shall, so far as possible, keep out of the way of a vessel restricted in her ability to maneuver.',
    "Q4": 'Both vessels alter course to starboard',
    "Q5": 'Seeing both sidelights of a vessel directly ahead',
    "Q6": 'Reduce your speed to the minimum at which it can be kept on course',
    "Q7": 'A vessel constrained by her draft',
    "Q8": 'Turn the vessel to starboard',
    "Q9": 'do not impair the visibility or distinctive character of the prescribed lights'
}

# Actual columns
question_map = {
    "Q1": 'The term "power-driven vessel" refers to any vessel __________.',
    "Q2": 'A vessel "restricted in her ability to maneuver" is one which __________.',
    "Q3": 'Which statement is TRUE, according to the Rules?',
    "Q4": 'When two power-driven vessels are meeting head-on and there is a risk of collision, which action is required to be taken?',
    "Q5": 'Which describes a head-on situation?',
    "Q6": 'Your vessel is underway in reduced visibility. You hear the fog signal of another vessel about 30° on your starboard bow. If danger of collision exists, which action(s) are you required to take?',
    "Q7": 'You have sighted three red lights in a vertical line on another vessel dead ahead at night. Which vessel would display these lights?',
    "Q8": 'By radar alone, you detect a vessel ahead on a collision course, about 3 miles distant. Your radar plot shows this to be a meeting situation. Which action should you take?',
    "Q9": 'A vessel may exhibit lights other than those prescribed by the Rules as long as the additional lights meet which requirement(s)?'
}

# Filter all-correct responders
cols = [question_map[q] for q in correct_answers]
df_knowledge = df[cols].astype(str).apply(lambda col: col.str.strip().str.lower())
correct_vals = {q: correct_answers[q].strip().lower() for q in correct_answers}
correct_mask = df_knowledge.apply(lambda row: all(row[question_map[q]] == correct_vals[q] for q in correct_answers), axis=1)
df = df[~correct_mask].copy()
print(f"Remaining responses after filters: {df.shape[0]}")

# ----------------------------------
# Analysis
# ----------------------------------
dimensions = ["Sufficiency", "Usefulness", "Clarity", "Adaptiveness", "Motivational Impact"]
feedback_types = ["Feedback A", "Feedback B", "Feedback C"]

ordinal_mapping_7 = {
    "Extremely Low": 1,
    "Very Low": 2,
    "Low": 3,
    "Moderate": 4,
    "High": 5,
    "Very High": 6,
    "Extremely High": 7
}

def get_column(feedback_type, dimension):
    for col in df.columns:
        if feedback_type in col and dimension in col:
            return col
    return None

# Friedman + Wilcoxon
friedman_results = []
wilcoxon_results = []
for dimension in dimensions:
    col_a = get_column("Feedback A", dimension)
    col_b = get_column("Feedback B", dimension)
    col_c = get_column("Feedback C", dimension)

    if col_a and col_b and col_c:
        data = df[[col_a, col_b, col_c]].apply(lambda col: col.map(ordinal_mapping_7.get)).dropna()
        data.columns = ["A", "B", "C"]
        f_stat, f_pval = friedmanchisquare(data["A"], data["B"], data["C"])
        friedman_results.append({"Metric": f"Friedman - {dimension}", "Statistic": round(f_stat, 3), "p-value": f"{round(f_pval, 4)}{' *' if f_pval < 0.05 else ''}", "Significant": f_pval < 0.05})
        for pair in [("A", "B"), ("A", "C"), ("B", "C")]:
            try:
                stat, p_val = wilcoxon(data[pair[0]], data[pair[1]])
                diffs = data[pair[0]] - data[pair[1]]
                median_diff = diffs.median()
                n_pos = (diffs > 0).sum()
                n_neg = (diffs < 0).sum()
                better = pair[0] if median_diff > 0 else pair[1] if median_diff < 0 else pair[0] if n_pos > n_neg else pair[1] if n_neg > n_pos else "Equal"
                wilcoxon_results.append({"Metric": f"Wilcoxon {pair[0]} vs {pair[1]} - {dimension}", "W-stat": round(stat, 3), "p-value": f"{round(p_val, 4)}{' *' if p_val < 0.05 else ''}", "Significant": p_val < 0.05, "Better": better})
            except ValueError:
                wilcoxon_results.append({"Metric": f"Wilcoxon {pair[0]} vs {pair[1]} - {dimension}", "W-stat": "NA", "p-value": "NA", "Significant": False, "Better": "NA"})

# Reliability
def cronbach_alpha(df_sub):
    df_sub = df_sub.dropna()
    item_scores = df_sub.values
    item_vars = item_scores.var(axis=0, ddof=1)
    total_var = item_scores.sum(axis=1).var(ddof=1)
    n_items = item_scores.shape[1]
    return np.nan if total_var == 0 else (n_items / (n_items - 1)) * (1 - item_vars.sum() / total_var)

encoder = OrdinalEncoder(categories=[list(ordinal_mapping_7.keys())] * len(dimensions))
reliability_scores = []
for fb_type in feedback_types:
    cols = [get_column(fb_type, dim) for dim in dimensions]
    df_sub = df[[col for col in cols if col]].dropna()
    df_encoded = pd.DataFrame(encoder.fit_transform(df_sub), columns=cols)
    alpha = cronbach_alpha(df_encoded)
    reliability_scores.append({"Feedback Type": fb_type, "Cronbach's Alpha": round(alpha, 3)})

# Means and STDs
mean_std_results = []
for dimension in dimensions:
    row = {"Dimension": dimension}
    for fb_type in feedback_types:
        col_name = get_column(fb_type, dimension)
        if col_name:
            values = df[col_name].dropna()
            numeric_values = values.map(ordinal_mapping_7)
            row[fb_type] = f"{numeric_values.mean():.2f} ± {numeric_values.std():.2f}"
    mean_std_results.append(row)

# Output
print("\n=== Friedman Test Results ===")
print(pd.DataFrame(friedman_results).to_string(index=False))
print("\n=== Wilcoxon Signed-Rank Test Results ===")
print(pd.DataFrame(wilcoxon_results).to_string(index=False))
print("\n=== Cronbach's Alpha ===")
print(pd.DataFrame(reliability_scores).to_string(index=False))
print("\n=== Mean ± Std Deviation ===")
print(pd.DataFrame(mean_std_results).to_string(index=False))
