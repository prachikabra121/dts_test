import pandas as pd

file_path = r"GLA_Service_Determination_Request_Monitoring.xlsx"

# Read both sheets
df_pacelogic = pd.read_excel(file_path, sheet_name="PACELOGIC")
df_smart = pd.read_excel(file_path, sheet_name="SMART SHEET DATA")

# Normalize column names (safer)
df_pacelogic.columns = df_pacelogic.columns.str.strip().str.lower()
df_smart.columns = df_smart.columns.str.strip().str.lower()

# Ensure 'title' exists
if "title" not in df_pacelogic.columns or "title" not in df_smart.columns:
    raise KeyError("⚠️ Column 'Title' not found in one of the sheets!")

# 1️⃣ Rows in SMART SHEET DATA not in PACELOGIC
not_found_in_pacelogic = df_smart[~df_smart["title"].isin(df_pacelogic["title"])]

# 2️⃣ Rows in PACELOGIC not in SMART SHEET DATA
not_found_in_smart = df_pacelogic[~df_pacelogic["title"].isin(df_smart["title"])]

# Save results into new sheets
with pd.ExcelWriter(file_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    not_found_in_pacelogic.to_excel(writer, sheet_name="NOT_FOUND_IN_PACELOGIC", index=False)
    not_found_in_smart.to_excel(writer, sheet_name="NOT_FOUND_IN_SMART", index=False)

print("✅ Comparison done. Missing rows written to 'NOT_FOUND_IN_PACELOGIC' and 'NOT_FOUND_IN_SMART'.")
