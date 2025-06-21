import pandas as pd
import requests
import time
from tqdm import tqdm
from datetime import datetime
from pytz import timezone

API_KEY = "aja9g1k52mr0co8l8h3csj64io1slb3j7t6se1ic5taar1skedktm3pa95qb5862"
TRACKING_ID = "pkmsalechanne-22"
INPUT_EXPORT = "KeepaExport-2025-ProductFinder.xlsx"
INPUT_TEMPLATE = "Final_Enriched_Affiliate_Template.xlsx"
OUTPUT_FILE = f"Final_Enriched_Affiliate_Sheet_{datetime.now(timezone('America/Vancouver')).strftime('%Y%m%d_%H%M')}.xlsx"

# --- Field Mappings ---
MAPPED_FIELDS = {
    "Title": "Title",
    "URL: Amazon": "URL: Amazon",
    "Sales Rank: Current": "Sales Rank",
    "Reviews: Rating": "Rating",
    "Reviews: Review Count": "Review Count",
    "Buy Box: Current": "Price (¥)",
    "Amazon: Buy Box: 1 day drop %": "Discount (%)",
    "Referral Fee %": "Referral Fee %",
    "Referral Fee based on current Buy Box price": "Referral Fee based on current Buy Box price",
    "Buy Box: Stock": "Buy Box: Stock",
    "One Time Coupon: Absolute": "One Time Coupon: Absolute",
    "One Time Coupon: Percentage": "One Time Coupon: Percentage",
    "One Time Coupon: Subscribe & Save %": "One Time Coupon: Subscribe & Save %",
    "ASIN": "ASIN"
}

# --- Load files ---
export_df = pd.read_excel(INPUT_EXPORT)
template_df = pd.read_excel(INPUT_TEMPLATE)

# --- Map fields ---
print("\n🔄 Mapping fields:")
for src, tgt in tqdm(MAPPED_FIELDS.items()):
    if src in export_df.columns and tgt in template_df.columns:
        template_df[tgt] = export_df[src]
    else:
        print(f"⚠️ Missing column: '{src}' in export or '{tgt}' in template")

# --- Logic: Extra Discount ---
def has_extra_discount(row):
    return "Yes" if any(pd.notna(row.get(col)) and str(row.get(col)).strip() != "" for col in [
        "One Time Coupon: Absolute", "One Time Coupon: Percentage", "One Time Coupon: Subscribe & Save %"
    ]) else ""
template_df["Extra Discount"] = template_df.apply(has_extra_discount, axis=1)

# --- API Fetch ---
asins = template_df["ASIN"].dropna().astype(str).tolist()
headers = {"Accept-Encoding": "gzip"}

print("\nQuerying Keepa API...")
all_products = []
for i in tqdm(range(0, len(asins), 10)):
    batch = asins[i:i + 10]
    url = f"https://api.keepa.com/product?key={API_KEY}&domain=5&buybox=1&stats=180&history=0&asin={','.join(batch)}"
    resp = requests.get(url, headers=headers)
    if resp.status_code == 200:
        all_products.extend(resp.json().get("products", []))
    time.sleep(2)

product_map = {p["asin"]: p for p in all_products}

# --- Logic injection ---
def enrich_row(row):
    asin = row["ASIN"]
    p = product_map.get(asin, {})
    stats = p.get("stats", {})

    # API fields
    row["isLowest"] = stats.get("isLowest")
    row["isLowest90"] = stats.get("isLowest90")
    row["buyBox30dDropPercent"] = stats.get("buyBox30dDropPercentage")

    # Coupon
    row["Coupon"] = p.get("coupon")

    # Seller Info
    row["Seller Info"] = p.get("sellerId")

    # Affiliate link
    row["Affiliate Link"] = f"https://www.amazon.co.jp/dp/{asin}/?tag={TRACKING_ID}"

    # Commission Tier (optional logic TBD)
    row["Commission Tier"] = "Tier ?"

    # Trending BSR
    sr30 = stats.get("salesRank30")
    sr90 = stats.get("salesRank90")
    if sr30 and sr90:
        row["Trending BSR"] = "↑ Improving" if sr30 < sr90 else ("↓ Worsening" if sr30 > sr90 else "→ Stable")
    else:
        row["Trending BSR"] = "N/A"

    # Variation Simplicity
    vc = p.get("variationCount")
    if vc is None:
        row["Variation Simplicity"] = "N/A"
    elif vc <= 3:
        row["Variation Simplicity"] = "✅ Simple"
    elif vc <= 10:
        row["Variation Simplicity"] = "⚠️ Moderate"
    else:
        row["Variation Simplicity"] = "❌ Complex"

    return row

# --- Apply enrichment ---
template_df = template_df.apply(enrich_row, axis=1)

# --- Save ---
template_df.to_excel(OUTPUT_FILE, index=False)
print(f"\n✅ Saved enriched file to {OUTPUT_FILE}")
