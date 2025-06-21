# Regenerate the full enrichment script with datetime import included
script_text = """
import pandas as pd
import requests
import os
from datetime import datetime
from tqdm import tqdm

# === CONFIGURATION ===
API_KEY = "aja9g1k52mr0co8l8h3csj64io1slb3j7t6se1ic5taar1skedktm3pa95qb5862"
AFFILIATE_TAG = "pkmsalechanne-22"
TEMPLATE_FILE = "Final_Enriched_Affiliate_Template.xlsx"
INPUT_FILE = "KeepaExport-2025-ProductFinder.xlsx"
TIMESTAMP = datetime.now().strftime("%Y%m%d_%H%M")
OUTPUT_FILE = f"Final_Enriched_Affiliate_Sheet_{TIMESTAMP}.xlsx"

# === LOAD DATA ===
print("Reading template and product finder...")
template_df = pd.read_excel(TEMPLATE_FILE)
product_df = pd.read_excel(INPUT_FILE)

# Standardize ASINs to string format
product_df['ASIN'] = product_df['ASIN'].astype(str)
template_df['ASIN'] = template_df['ASIN'].astype(str)

# === MERGE KNOWN FIELDS ===
print("Merging known product fields...")
merged_df = template_df.copy()
merged_df = merged_df.drop(columns=[col for col in merged_df.columns if col != 'ASIN'])
merged_df = pd.merge(merged_df, product_df[[
    'ASIN',
    'Title',
    'URL: Amazon',
    'Sales Rank: Current',
    'Reviews: Rating',
    'Reviews: Review Count',
    'Buy Box: Current',
    'Referral Fee %',
    'Referral Fee based on current Buy Box price',
    'Buy Box: Stock'
]], on='ASIN', how='left')

# Rename to match final sheet headers
merged_df.rename(columns={
    'Title': 'Title',
    'URL: Amazon': 'URL: Amazon',
    'Sales Rank: Current': 'Sales Rank',
    'Reviews: Rating': 'Rating',
    'Reviews: Review Count': 'Review Count',
    'Buy Box: Current': 'Price (¥)',
    'Referral Fee %': 'Referral Fee %',
    'Referral Fee based on current Buy Box price': 'Referral Fee based on current Buy Box price',
    'Buy Box: Stock': 'Buy Box: Stock'
}, inplace=True)

# === GENERATE AFFILIATE LINK ===
print("Generating affiliate links...")
merged_df["Affiliate Link"] = merged_df["ASIN"].apply(
    lambda asin: f"https://www.amazon.co.jp/dp/{asin}/?tag={AFFILIATE_TAG}"
)

# === PLACEHOLDERS FOR API-ENRICHED COLUMNS ===
api_fields = [
    "Discount (%)", "isLowest", "isLowest90", "buyBox30dDropPercent", "Coupon",
    "Grade", "Fake Drop Flag", "Seller Info", "Commission Tier",
    "Trending BSR", "Variation Simplicity", "Marketing Approved",
    "20-sec Pitch", "Promo Image Placeholder", "Image", "Chart"
]
for col in api_fields:
    if col not in merged_df.columns:
        merged_df[col] = None

# Reorder columns to match template
final_columns = pd.read_excel(TEMPLATE_FILE, nrows=0).columns.tolist()
merged_df = merged_df.reindex(columns=final_columns)

# === SAVE FILE ===
merged_df.to_excel(OUTPUT_FILE, index=False)
print(f"✅ Enrichment template prepared: {OUTPUT_FILE}")
"""

# Save the script locally
script_path = "/mnt/data/enrich_affiliate_sheet.py"
with open(script_path, "w", encoding="utf-8") as f:
    f.write(script_text)

script_path
