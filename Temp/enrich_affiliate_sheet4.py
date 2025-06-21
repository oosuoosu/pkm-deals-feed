import pandas as pd
from openpyxl import load_workbook
from datetime import datetime

# === FILE PATHS ===
INPUT_FILE = "KeepaExport-2025-ProductFinder.xlsx"
TEMPLATE_FILE = "Final_Enriched_Affiliate_Template.xlsx"
OUTPUT_FILE = f"Final_Enriched_Affiliate_Sheet_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

# === LOAD DATA ===
df_input = pd.read_excel(INPUT_FILE)
df_template = pd.read_excel(TEMPLATE_FILE)

# === CLEAN UP COLUMN NAMES ===
df_input.columns = df_input.columns.str.strip()
df_template.columns = df_template.columns.str.strip()

# === MAP FIELDS BY ASIN ===
mapping = {
    "Title": "Title",
    "URL: Amazon": "URL: Amazon",
    "Sales Rank: Current": "Sales Rank",
    "Reviews: Rating": "Rating",
    "Reviews: Review Count": "Review Count",
    "Buy Box: Current": "Price (¥)",
    "Referral Fee %": "Referral Fee %",
    "Referral Fee based on current Buy Box price": "Referral Fee based on current Buy Box price",
    "Buy Box: Stock": "Buy Box: Stock",
}

# === ENSURE ASIN COLUMN EXISTS IN BOTH ===
if "ASIN" not in df_input.columns or "ASIN" not in df_template.columns:
    raise ValueError("Missing 'ASIN' column in input or template sheet.")

df_input.set_index("ASIN", inplace=True)
df_template.set_index("ASIN", inplace=True)

# === APPLY FIELD MAPPINGS ===
for input_col, target_col in mapping.items():
    if input_col in df_input.columns and target_col in df_template.columns:
        df_template[target_col] = df_input[input_col]

# === RESET INDEX TO SAVE ASIN BACK TO FILE ===
df_template.reset_index(inplace=True)

# === SAVE OUTPUT ===
df_template.to_excel(OUTPUT_FILE, index=False)
print(f"\n✅ Enriched file saved as: {OUTPUT_FILE}")
