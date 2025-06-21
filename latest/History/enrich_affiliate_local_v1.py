from datetime import datetime
import pandas as pd
import os
from tqdm import tqdm

# File paths
INPUT_EXPORT = "KeepaExport-2025-ProductFinder.xlsx"
TEMPLATE_FILE = "Final_Enriched_Affiliate_Template.xlsx"
TIMESTAMP = datetime.now().strftime("%Y%m%d_%H%M")
OUTPUT_FILE = f"Final_Enriched_Affiliate_Sheet_{TIMESTAMP}.xlsx"

# Tracking ID for affiliate link
AFFILIATE_TAG = "pkmsalechanne-22"

# Column mappings
MAPPED_COLUMNS = {
    "Title": "Title",
    "URL: Amazon": "URL: Amazon",
    "Sales Rank: Current": "Sales Rank",
    "Reviews: Rating": "Rating",
    "Reviews: Review Count": "Review Count",
    "Buy Box: Current": "Price (¥)",
    "Buy Box: 1 day drop %": "Discount (%)",
    "One Time Coupon: Absolute": "One Time Coupon: Absolute",
    "One Time Coupon: Percentage": "One Time Coupon: Percentage",
    "One Time Coupon: Subscribe & Save %": "One Time Coupon: Subscribe & Save %",
    "Buy Box: Stock": "Buy Box: Stock",
    "Referral Fee %": "Referral Fee %",
    "Referral Fee based on current Buy Box price": "Referral Fee based on current Buy Box price (¥)",
    "ASIN": "ASIN"
}

# Enrichment logic
def enrich_logic(row):
    asin = row.get("ASIN")
    row["Affiliate Link"] = f"https://www.amazon.co.jp/dp/{asin}/?tag={AFFILIATE_TAG}" if asin else ""

    if any([
        pd.notna(row.get("One Time Coupon: Absolute")),
        pd.notna(row.get("One Time Coupon: Percentage")),
        pd.notna(row.get("One Time Coupon: Subscribe & Save %"))
    ]):
        row["Extra Discount"] = "Yes"
    else:
        row["Extra Discount"] = ""

    try:
        discount = float(row.get("Discount (%)", 0))
    except:
        discount = 0
    is_lowest_90 = row.get("isLowest90", True)
    row["Fake Drop Flag"] = not is_lowest_90 and discount >= 30

    return row

# Load both files
export_df = pd.read_excel(INPUT_EXPORT)
template_df = pd.read_excel(TEMPLATE_FILE)

# Map columns
for src, tgt in tqdm(MAPPED_COLUMNS.items(), desc="Mapping columns"):
    if src in export_df.columns and tgt in template_df.columns:
        template_df[tgt] = export_df[src]
    else:
        print(f"⚠️ Missing column: '{src}' in export or '{tgt}' in template")

# Enrich logic
template_df = template_df.apply(enrich_logic, axis=1)

# Save
template_df.to_excel(OUTPUT_FILE, index=False)
print(f"✅ Saved enriched file to {OUTPUT_FILE}")
