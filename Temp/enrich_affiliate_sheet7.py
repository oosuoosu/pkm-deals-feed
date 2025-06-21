import pandas as pd
import requests
import time
from datetime import datetime
from tqdm import tqdm

# ===== Configuration =====
API_KEY = "aja9g1k52mr0co8l8h3csj64io1slb3j7t6se1ic5taar1skedktm3pa95qb5862"
TRACKING_ID = "pkmsalechanne-22"
INPUT_FILE = "KeepaExport-2025-ProductFinder.xlsx"
TEMPLATE_FILE = "Final_Enriched_Affiliate_Template.xlsx"
TIMESTAMP = datetime.now().strftime("%Y%m%d_%H%M")
OUTPUT_FILE = f"Final_Enriched_Affiliate_Sheet_{TIMESTAMP}.xlsx"

# ===== Column Mapping from Input to Template =====
FIELD_MAPPING = {
    "Title": "Title",
    "URL: Amazon": "URL: Amazon",
    "Sales Rank: Current": "Sales Rank",
    "Reviews: Rating": "Rating",
    "Reviews: Review Count": "Review Count",
    "Buy Box: Current": "Price (¥)",
    "Referral Fee %": "Referral Fee %",
    "Referral Fee based on current Buy Box price": "Referral Fee based on current Buy Box price",
    "Buy Box: Stock": "Buy Box: Stock",
    "ASIN": "ASIN"
}

# ===== Logic Functions =====
def compute_grade(discount, sales_rank):
    if pd.isna(discount) or pd.isna(sales_rank):
        return "F"
    if discount >= 50 and sales_rank <= 5000:
        return "A"
    elif discount >= 40 and sales_rank <= 10000:
        return "B"
    elif discount >= 30 and sales_rank <= 20000:
        return "C"
    elif discount >= 20 or sales_rank <= 50000:
        return "D"
    return "F"

def compute_fake_drop(stats):
    return not stats.get("isLowest90", False) and stats.get("buyBox30dDropPercent", 0) >= 30

def compute_trending_bsr(stats):
    if stats.get("salesRank30") and stats.get("salesRank90"):
        if stats["salesRank30"] < stats["salesRank90"]:
            return "↑ Improving"
        elif stats["salesRank30"] > stats["salesRank90"]:
            return "↓ Worsening"
        else:
            return "→ Stable"
    return "N/A"

def compute_variation_simplicity(product):
    count = product.get("variationCount")
    if count is None:
        return "N/A"
    elif count <= 3:
        return "✅ Simple"
    elif count <= 10:
        return "⚠️ Moderate"
    else:
        return "❌ Complex"

def fetch_keepa_data(asin):
    url = f"https://api.keepa.com/product?key={API_KEY}&domain=5&asin={asin}&stats=180"
    r = requests.get(url)
    if r.status_code != 200:
        return None
    data = r.json()
    if not data.get("products"):
        return None
    return data["products"][0]

# ===== Main Enrichment =====
def enrich():
    input_df = pd.read_excel(INPUT_FILE)
    template_df = pd.read_excel(TEMPLATE_FILE)
    enriched_df = template_df.copy()

    asin_list = input_df["ASIN"].dropna().astype(str).tolist()

    for idx, asin in enumerate(tqdm(asin_list, desc="Enriching ASINs")):
        # Transfer fields from ProductFinder
        source_row = input_df[input_df["ASIN"] == asin]
        if source_row.empty:
            continue

        for source_col, target_col in FIELD_MAPPING.items():
            if source_col in source_row.columns:
                enriched_df.at[idx, target_col] = source_row.iloc[0][source_col]

        enriched_df.at[idx, "ASIN"] = asin
        enriched_df.at[idx, "Affiliate Link"] = f"https://www.amazon.co.jp/dp/{asin}/?tag={TRACKING_ID}"

        # Fetch Keepa data
        product = fetch_keepa_data(asin)
        if not product:
            continue

        stats = product.get("stats", {})

        # Discount Calculation
        buy_box_price = stats.get("buyBoxPrice", [None])[-1]
        buy_box_avg = stats.get("buyBox30", None)
        discount = None
        if buy_box_price and buy_box_avg and buy_box_avg > 0:
            discount = round((1 - buy_box_price / buy_box_avg) * 100, 2)
        enriched_df.at[idx, "Discount (%)"] = discount

        # Keepa Stats
        enriched_df.at[idx, "isLowest"] = stats.get("isLowest")
        enriched_df.at[idx, "isLowest90"] = stats.get("isLowest90")
        enriched_df.at[idx, "buyBox30dDropPercent"] = stats.get("buyBox30dDropPercent")
        enriched_df.at[idx, "Coupon"] = product.get("coupon", {}).get("couponAmount")
        enriched_df.at[idx, "Seller Info"] = product.get("sellerId")

        # Calculated Fields
        rank_val = source_row.iloc[0].get("Sales Rank: Current", None)
        enriched_df.at[idx, "Grade"] = compute_grade(discount, rank_val)
        enriched_df.at[idx, "Fake Drop Flag"] = compute_fake_drop(stats)
        enriched_df.at[idx, "Commission Tier"] = "Standard"
        enriched_df.at[idx, "Trending BSR"] = compute_trending_bsr(stats)
        enriched_df.at[idx, "Variation Simplicity"] = compute_variation_simplicity(product)

        # Token recovery pause
        if (idx + 1) % 2 == 0:
            time.sleep(6)

    enriched_df.to_excel(OUTPUT_FILE, index=False)
    print(f"\n✅ Done! Enriched sheet saved as: {OUTPUT_FILE}")

# ===== Run It =====
if __name__ == "__main__":
    enrich()
