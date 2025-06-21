import time
import pandas as pd
import requests
from tqdm import tqdm

# === CONFIGURATION ===
API_KEY = "aja9g1k52mr0co8l8h3csj64io1slb3j7t6se1ic5taar1skedktm3pa95qb5862"
INPUT_FILE = "KeepaExport-2025-ProductFinder.xlsx"
OUTPUT_FILE = "EnrichedAffiliateSheet_Test10_FIXED.xlsx"
STARTING_TOKENS = 1260
TOKENS_PER_REQUEST = 10  # As each ASIN costs about 10 tokens
RECOVERY_RATE_PER_MIN = 20  # tokens
RECOVERY_INTERVAL_SEC = 3  # one token every 3 seconds

# === FUNCTIONS ===
def get_asins_from_excel(file_path, limit=10):
    df = pd.read_excel(file_path)
    return df['ASIN'].dropna().astype(str).tolist()[:limit]

def get_commission_tier(referral_fee_pct):
    if referral_fee_pct is None:
        return "Tier 4"
    if referral_fee_pct >= 12:
        return "Tier 1"
    elif referral_fee_pct >= 8:
        return "Tier 2"
    elif referral_fee_pct >= 5:
        return "Tier 3"
    else:
        return "Tier 4"

def extract_product_info(product):
    asin = product.get('asin')
    title = product.get('title', '')
    url = f"https://www.amazon.co.jp/dp/{asin}"
    
    stats = product.get('stats', {})
    if isinstance(stats.get("buyBoxPrice"), list):
        buy_box_price = stats.get("buyBoxPrice", [None])[-1]
    else:
        buy_box_price = stats.get("buyBoxPrice")

    referral_fee_pct = product.get("referralFee")
    commission_tier = get_commission_tier(referral_fee_pct)

    return {
        "ASIN": asin,
        "Title": title,
        "Buy Box Price (\u00a5)": buy_box_price / 100 if buy_box_price else None,
        "Referral Fee %": referral_fee_pct,
        "Commission Tier": commission_tier,
        "URL": url
    }

def fetch_keepa_data(asins):
    asin_str = ",".join(asins)
    url = f"https://api.keepa.com/product?key={API_KEY}&domain=5&buybox=1&stats=180&history=0&asin={asin_str}"
    response = requests.get(url)
    if response.status_code == 200:
        return response.json().get("products", [])
    else:
        print("Error from Keepa API:", response.text)
        return []

# === MAIN EXECUTION ===
all_asins = get_asins_from_excel(INPUT_FILE, limit=10)
remaining_tokens = STARTING_TOKENS
batch_size = TOKENS_PER_REQUEST  # 10 tokens per product

results = []
print("Starting batch processing with token recovery strategy...")

for asin in tqdm(all_asins, desc="Processing ASINs"):
    if remaining_tokens < TOKENS_PER_REQUEST:
        print("\nWaiting for token recovery...")
        while remaining_tokens < TOKENS_PER_REQUEST:
            time.sleep(RECOVERY_INTERVAL_SEC)
            remaining_tokens += 1

    data = fetch_keepa_data([asin])
    remaining_tokens -= TOKENS_PER_REQUEST

    if data:
        results.append(extract_product_info(data[0]))
    else:
        print(f"Failed to fetch data for ASIN: {asin}")

# === SAVE OUTPUT ===
df = pd.DataFrame(results)
df.to_excel(OUTPUT_FILE, index=False)
print(f"\n✅ Done. Results saved to {OUTPUT_FILE}")