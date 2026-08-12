import pandas as pd
import os

# Define file paths based on your folder setup
ifrs_csv = "Template (Silvassa) revised _ Asset Master_Source file.xlsx - IFRS.csv"
igaap_csv = "Template (Silvassa) revised _ Asset Master_Source file.xlsx - I GAAP.csv"
output_excel = "Combined_Asset_Source.xlsx"

print("🔄 Reading split CSV data...")

if not os.path.exists(ifrs_csv) or not os.path.exists(igaap_csv):
    print("❌ Error: Please ensure both CSV files are in this exact folder and named correctly!")
else:
    # Read the CSVs (loading all columns as strings to preserve leading zeros in asset numbers)
    df_ifrs = pd.read_csv(ifrs_csv, dtype=str)
    df_igaap = pd.read_csv(igaap_csv, dtype=str)

    print("📊 Merging into single Excel Workbook with separate tabs...")
    with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
        df_ifrs.to_excel(writer, sheet_name="IFRS", index=False)
        df_igaap.to_excel(writer, sheet_name="IGAAP", index=False)

    print(f"✅ Success! Created unified file: {output_excel}")