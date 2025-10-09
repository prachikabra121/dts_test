

import pandas as pd

# File path
file_path = "GLA_Service_Determination_Request_Monitoring.xlsx"   # change this to your actual Excel file

# Read only the sheet "SMART SHEET DATA"
df = pd.read_excel(file_path, sheet_name="PACELOGIC")

# Ensure date is properly formatted
df["mph_original_requested_date"] = pd.to_datetime(df["mph_original_requested_date"], errors="coerce")

# Create new column combining first name + last name + date
df["Title"] = (
    df["firstname"].astype(str) + "_" +
    df["lastname"].astype(str) + "_" +
    df["mph_original_requested_date"].dt.strftime("%Y-%m-%d")
)

# Write back to Excel, preserving other sheets if they exist
with pd.ExcelWriter(file_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    df.to_excel(writer, sheet_name="PACELOGIC", index=False)

print("✅ New column 'Full Info' added to sheet 'SMART SHEET DATA'")
