import pandas as pd
import warnings
import numpy as np

# Suppress openpyxl warnings
warnings.simplefilter("ignore", UserWarning)

file_path = "tes.xlsx"  # <-- change this to your file

# ---- Load sheets ----
new_df = pd.read_excel(file_path, sheet_name="new")  # New data
old_df = pd.read_excel(file_path, sheet_name="old")  # Old data


# ---- Normalize keys ----
def normalize_key(series):
    return (
        series.astype(str)
        .str.strip()
        .str.lower()
        .replace({"nan": "", "null": ""})  # clean NULLs
    )


def clean_adlshiftid(series):
    s = series.astype(str).str.strip().str.lower()
    s = s.str.replace(r":00\.0$", "", regex=True)  # convert "26:00.0" → "26"
    s = s.str.replace(r"\.0$", "", regex=True)  # convert "39.0" → "39"
    s = s.replace({"nan": "", "null": ""})
    return s


# Apply cleaning
if "memberid" in new_df.columns: new_df["memberid"] = normalize_key(new_df["memberid"])
if "memberid" in old_df.columns: old_df["memberid"] = normalize_key(old_df["memberid"])

if "adlshiftid" in new_df.columns: new_df["adlshiftid"] = clean_adlshiftid(new_df["adlshiftid"])
if "adlshiftid" in old_df.columns: old_df["adlshiftid"] = clean_adlshiftid(old_df["adlshiftid"])

# ---- Define key ----
key_cols = ["adlshiftid"]

# ---- Debug key overlap ----
common_keys = pd.merge(old_df[key_cols], new_df[key_cols], how="inner").drop_duplicates()
print("🔑 New keys:", new_df[key_cols].drop_duplicates().shape[0])
print("🔑 Old keys:", old_df[key_cols].drop_duplicates().shape[0])
print("🔑 Common keys:", common_keys.shape[0])
print("Sample common keys:\n", common_keys.head(10))

# ---- Merge old vs new ----
merged = pd.merge(
    old_df, new_df,
    on=key_cols,
    how="outer",
    suffixes=("_old", "_new"),
    indicator=True
)


# ---- Cleaning function ----
def clean_text(val):
    """Normalize string, numeric, or datetime values for comparison."""
    if pd.isna(val):
        return np.nan
    if isinstance(val, str):
        return val.strip().lower()
    return val  # keep numeric/datetime as-is


# Columns to check for changes (excluding keys)
check_cols = [
    "msemr_azurefhirid", "memberid", "mph_date_time", "mph_bath", "mph_bed_mobility", "mph_behaviors",
    "mph_bladder", "mph_bowel_movement_size", "mph_bowels",
    "mph_dressing_updated", "mph_meal_intake", "mph_mobility_activity",
    "mph_safety_fall", "mph_skin_observation", "mph_transfer",
    "shift", "comments", "modifiedby"
]

for col in check_cols:
    if f"{col}_old" in merged.columns:
        merged[f"{col}_old_clean"] = merged[f"{col}_old"].apply(clean_text)
    if f"{col}_new" in merged.columns:
        merged[f"{col}_new_clean"] = merged[f"{col}_new"].apply(clean_text)


# ---- Classification ----
def classify(row):
    if row["_merge"] == "left_only":
        return "old"
    elif row["_merge"] == "right_only":
        return "new"
    else:
        # Compare cleaned columns
        for col in check_cols:
            old_val = row.get(f"{col}_old_clean")
            new_val = row.get(f"{col}_new_clean")

            # Treat NaN as equal
            if pd.isna(old_val) and pd.isna(new_val):
                continue
            if old_val != new_val:
                return "updated"
        return "unchanged"


merged["flag"] = merged.apply(classify, axis=1)

# ---- Build final result ----
final = pd.DataFrame()
for col in key_cols + check_cols + ["created"]:
    if f"{col}_new" in merged.columns and f"{col}_old" in merged.columns:
        final[col] = merged[f"{col}_new"].combine_first(merged[f"{col}_old"])
    elif f"{col}_new" in merged.columns:
        final[col] = merged[f"{col}_new"]
    elif f"{col}_old" in merged.columns:
        final[col] = merged[f"{col}_old"]

final["flag"] = merged["flag"]

# ---- Save results ----
# Option 1: Only NEW & UPDATED rows
final_filtered = final[final["flag"].isin(["new", "updated"])]
final_filtered.to_excel("comparison_result_tes.xlsx", index=False)

# Option 2: All rows including unchanged
final.to_excel("comparison_result_tes_all.xlsx", index=False)

print(f"✅ Comparison complete!")
print(f"Saved NEW & UPDATED rows to: comparison_result_tes.xlsx")
print(f"Saved ALL rows including unchanged to: comparison_result_tes_all.xlsx")
