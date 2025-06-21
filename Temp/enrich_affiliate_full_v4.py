from datetime import datetime
import pandas as pd
from keepa import Keepa
from tqdm import tqdm

# --- CONFIGURATION ---
API_KEY = "aja9g1k52mr0co8l8h3csj64io1slb3j7t6se1ic5taar1skedktm3pa95qb5862"
INPUT_EXPORT = "KeepaExport-2025-ProductFinder.xlsx"
TEMPLATE_FILE = "Final_Enriched_Affiliate_Template.xlsx"
TIMESTAMP = datetime.now().strftime("%Y%m%d_%H%M")
OUTPUT_FILE = f"Final_Enriched_Affiliate_Sheet_{TIMESTAMP}.xlsx"
AFFILIATE_TAG = "pkmsalechanne-22"

# --- COLUMN MAPPINGS ---
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
    "ASIN": "ASIN",
}

# --- LOGIC HELPERS ---
def enrich_row_local(row):
    asin = row.get("ASIN", "")
    row["Affiliate Link"] = f"https://www.amazon.co.jp/dp/{asin}/?tag={AFFILIATE_TAG}" if asin else ""

    if any([
        pd.notna(row.get("One Time Coupon: Absolute")),
        pd.notna(row.get("One Time Coupon: Percentage")),
        pd.notna(row.get("One Time Coupon: Subscribe & Save %"))
    ]):
        row["Extra Discount"] = "Yes"
    else:
        row["Extra Discount"] = ""

    return row

def safe_bool(val):
    if isinstance(val, list):
        return bool(val[-1])
    return bool(val)

def safe_float(val):
    if isinstance(val, list):
        return float(val[-1])
    return float(val) if val is not None else ""

# --- LOAD FILES ---
export_df = pd.read_excel(INPUT_EXPORT)
template_df = pd.read_excel(TEMPLATE_FILE)

# --- COLUMN MAPPING ---
print("🔄 Mapping local columns:")
for src, tgt in tqdm(MAPPED_COLUMNS.items(), desc="Mapping columns"):
    if src in export_df.columns and tgt in template_df.columns:
        template_df[tgt] = export_df[src]
    else:
        print(f"⚠️ Missing column: '{src}' in export or '{tgt}' in template")

# --- APPLY LOCAL ENRICHMENT ---
template_df = template_df.apply(enrich_row_local, axis=1)

# --- KEEP API ENRICHMENT ---
print("🔌 Querying Keepa API...")
asins = template_df["ASIN"].dropna().astype(str).tolist()
keepa = Keepa(API_KEY)
products = keepa.query(asins, domain="JP", stats=180)

print("⚙️ Enriching via Keepa API...")
for idx, asin in enumerate(tqdm(asins, desc="Enriching ASINs")):
    product = next((p for p in products if p.get("asin") == asin), {})
    stats = product.get("stats", {})

    try:
        template_df.at[idx, "isLowest"] = safe_bool(stats.get("isLowest", False))
        template_df.at[idx, "isLowest90"] = safe_bool(stats.get("isLowest90", False))
        template_df.at[idx, "buyBox30dDropPercent"] = safe_float(stats.get("buyBox30dDropPercentage"))

        # Trending BSR
        sr30 = stats.get("salesRank30")
        sr90 = stats.get("salesRank90")
        if sr30 and sr90:
            trend = "↑ Improving" if sr30 < sr90 else "↓ Worsening" if sr30 > sr90 else "→ Stable"
        else:
            trend = "N/A"
        template_df.at[idx, "Trending BSR"] = trend

        # Variation Simplicity
        var_count = product.get("variationCount")
        if var_count is not None:
            if var_count <= 3:
                var_label = "✅ Simple"
            elif var_count <= 10:
                var_label = "⚠️ Moderate"
            else:
                var_label = "❌ Complex"
        else:
            var_label = "N/A"
        template_df.at[idx, "Variation Simplicity"] = var_label

        # Fake Drop Flag
        is_l90 = safe_bool(stats.get("isLowest90", False))
        drop_pct = safe_float(stats.get("buyBox30dDropPercentage"))
        fake_flag = not is_l90 and drop_pct >= 30 if drop_pct != "" else False
        template_df.at[idx, "Fake Drop Flag"] = fake_flag

        # Grade logic
        discount = template_df.at[idx, "Discount (%)"]
        rank = template_df.at[idx, "Sales Rank"]
        grade = "F"
        if pd.notna(discount) and pd.notna(rank):
            try:
                discount = float(discount)
                rank = int(rank)
                if discount >= 50 and rank <= 5000:
                    grade = "A"
                elif discount >= 40 and rank <= 10000:
                    grade = "B"
                elif discount >= 30 and rank <= 20000:
                    grade = "C"
                elif discount >= 20 or rank <= 50000:
                    grade = "D"
            except:
                grade = "F"
        if fake_flag and grade in "ABCD":
            grade = {"A": "B", "B": "C", "C": "D", "D": "F"}[grade]
        template_df.at[idx, "Grade"] = grade

    except Exception as e:
        print(f"❌ Error enriching ASIN {asin}: {e}")

# --- SAVE FILE ---
template_df.to_excel(OUTPUT_FILE, index=False)
print(f"✅ Saved enriched file to {OUTPUT_FILE}")
