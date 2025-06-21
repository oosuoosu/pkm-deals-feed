import pandas as pd
import requests
import time
from datetime import datetime
from tqdm import tqdm

# --- Configuration ---
API_KEY = "aja9g1k52mr0co8l8h3csj64io1slb3j7t6se1ic5taar1skedktm3pa95qb5862"
INPUT_FILE = "KeepaExport-2025-ProductFinder.xlsx"
TEMPLATE_FILE = "Final_Enriched_Affiliate_Template.xlsx"
OUTPUT_TIMESTAMP = datetime.now().strftime("%Y%m%d_%H%M")
OUTPUT_FILE = f"Final_Enriched_Affiliate_Sheet_{OUTPUT_TIMESTAMP}.xlsx"
TRACKING_ID = "pkmsalechanne-22"

# --- Utility Functions ---
def fetch_keepa_data(asin, retry=3, delay=5):
    for attempt in range(retry):
        try:
            url = f"https://api.keepa.com/product?key={API_KEY}&domain=5&buybox=1&stats=180&asin={asin}"
            response = requests.get(url, timeout=10)
            if response.status_code == 200:
                return response.json()
            else:
                print(f"Error fetching {asin}: {response.status_code}")
                return None
        except Exception as e:
            print(f"Attempt {attempt+1} failed for {asin}: {e}")
            time.sleep(delay)
    return None

def compute_grade(discount, rank):
    if pd.isna(discount) or pd.isna(rank):
        return "F"
    if discount >= 50 and rank <= 5000:
        return "A"
    elif discount >= 40 and rank <= 10000:
        return "B"
    elif discount >= 30 and rank <= 20000:
        return "C"
    elif discount >= 20 or rank <= 50000:
        return "D"
    else:
        return "F"

def compute_fake_drop_flag(product_data):
    try:
        stats = product_data["products"][0].get("stats", {})
        is_lowest = stats.get("isLowest", False)
        is_lowest90 = stats.get("isLowest90", False)
        drop_30d = stats.get("buyBox30dDropPercent", 0)
        return not is_lowest90 and drop_30d >= 30
    except:
        return False

def compute_trending_bsr(stats):
    try:
        if stats.get("salesRank30") and stats.get("salesRank90"):
            if stats["salesRank30"] < stats["salesRank90"]:
                return "↑ Improving"
            elif stats["salesRank30"] > stats["salesRank90"]:
                return "↓ Worsening"
            else:
                return "→ Stable"
        else:
            return "N/A"
    except:
        return "N/A"

def compute_variation_simplicity(product_data):
    vc = product_data["products"][0].get("variationCount", None)
    if vc is None:
        return "N/A"
    elif vc <= 3:
        return "✅ Simple"
    elif vc <= 10:
        return "⚠️ Moderate"
    else:
        return "❌ Complex"

def extract_affiliate_link(asin):
    return f"https://www.amazon.co.jp/dp/{asin}/?tag={TRACKING_ID}"

# --- Main Logic ---
def enrich():
    print("Loading files...")
    product_df = pd.read_excel(INPUT_FILE)
    enriched_df = pd.read_excel(TEMPLATE_FILE)

    enriched_df["ASIN"] = product_df["ASIN"]

    # Mapped Fields (overwrite if exists)
    field_map = {
        "Title": "Title",
        "URL: Amazon": "URL: Amazon",
        "Sales Rank: Current": "Sales Rank",
        "Reviews: Rating": "Rating",
        "Reviews: Review Count": "Review Count",
        "Buy Box: Current": "Price (¥)",
        "Referral Fee %": "Referral Fee %",
        "Referral Fee based on current Buy Box price": "Referral Fee based on current Buy Box price",
        "Buy Box: Stock": "Buy Box: Stock"
    }

    for src, dst in field_map.items():
        if src in product_df.columns and dst in enriched_df.columns:
            enriched_df[dst] = product_df[src]

    print("Enriching via Keepa API...")
    for idx in tqdm(enriched_df.index, desc="Enriching ASINs"):
        asin = enriched_df.at[idx, "ASIN"]
        result = fetch_keepa_data(asin)
        if not result or not result.get("products"):
            continue
        stats = result["products"][0].get("stats", {})
        enriched_df.at[idx, "Discount (%)"] = stats.get("buyBoxPriceDropPercent", None)
        enriched_df.at[idx, "isLowest"] = stats.get("isLowest", False)
        enriched_df.at[idx, "isLowest90"] = stats.get("isLowest90", False)
        enriched_df.at[idx, "buyBox30dDropPercent"] = stats.get("buyBox30dDropPercent", None)
        enriched_df.at[idx, "Coupon"] = result["products"][0].get("coupon", None)
        enriched_df.at[idx, "Seller Info"] = result["products"][0].get("sellerId", "")
        enriched_df.at[idx, "Affiliate Link"] = extract_affiliate_link(asin)
        enriched_df.at[idx, "Grade"] = compute_grade(enriched_df.at[idx, "Discount (%)"], enriched_df.at[idx, "Sales Rank"])
        enriched_df.at[idx, "Fake Drop Flag"] = compute_fake_drop_flag(result)
        enriched_df.at[idx, "Commission Tier"] = "Standard"
        enriched_df.at[idx, "Trending BSR"] = compute_trending_bsr(stats)
        enriched_df.at[idx, "Variation Simplicity"] = compute_variation_simplicity(result)

    print(f"Saving to {OUTPUT_FILE}")
    enriched_df.to_excel(OUTPUT_FILE, index=False)
    print("Done.")

# --- Entry Point ---
if __name__ == "__main__":
    enrich()
