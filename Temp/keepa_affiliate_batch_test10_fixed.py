
import time
import requests
import pandas as pd
from tqdm import tqdm

# Load ASINs from Excel
df = pd.read_excel("KeepaExport-2025-ProductFinder.xlsx")
asins = df["ASIN"].dropna().astype(str).tolist()[:10]  # First 10 ASINs

# Your Keepa API token
API_TOKEN = "aja9g1k52mr0co8l8h3csj64io1slb3j7t6se1ic5taar1skedktm3pa95qb5862"

# Token bucket settings
MAX_TOKENS = 1260
TOKENS_PER_MIN = 20
TOKENS_USED = 0
LAST_REFILL = time.time()

# Function to simulate token refill
def refill_tokens():
    global TOKENS_USED, LAST_REFILL
    now = time.time()
    recovered = int((now - LAST_REFILL) / 3)  # 1 token per 3 sec
    TOKENS_USED = max(0, TOKENS_USED - recovered)
    LAST_REFILL = now

# Extract product info
def extract_product_info(data):
    asin = data.get("asin", "")
    title = data.get("title", "")
    url = f"https://www.amazon.co.jp/dp/{asin}"

    stats = data.get("stats", {})
    buy_box_price = None
    if isinstance(stats.get("buyBoxPrice"), list) and stats["buyBoxPrice"]:
        buy_box_price = stats["buyBoxPrice"][-1] / 100 if stats["buyBoxPrice"][-1] is not None else None

    referral_fee_percent = data.get("referralFeePercent", None)

    # Commission Tier
    if referral_fee_percent is not None:
        if referral_fee_percent >= 12:
            tier = "Tier 1"
        elif referral_fee_percent >= 8:
            tier = "Tier 2"
        elif referral_fee_percent >= 5:
            tier = "Tier 3"
        else:
            tier = "Tier 4"
    else:
        tier = "Tier 4"

    return {
        "ASIN": asin,
        "Title": title,
        "Buy Box Price (¥)": buy_box_price,
        "Referral Fee %": referral_fee_percent,
        "Commission Tier": tier,
        "URL": url
    }

# Output list
results = []

# Fetch product info from Keepa
for asin in tqdm(asins, desc="Processing ASINs"):
    refill_tokens()
    if TOKENS_USED >= MAX_TOKENS:
        sleep_time = 3
        print(f"Rate limit hit. Sleeping for {sleep_time}s...")
        time.sleep(sleep_time)
        refill_tokens()

    url = f"https://api.keepa.com/product?key={API_TOKEN}&domain=5&asin={asin}&buybox=1&stats=180"
    try:
        response = requests.get(url)
        TOKENS_USED += 1
        data = response.json()
        if "products" in data and data["products"]:
            product_info = extract_product_info(data["products"][0])
            results.append(product_info)
    except Exception as e:
        print(f"Error processing ASIN {asin}: {e}")

# Save to Excel
output_df = pd.DataFrame(results)
output_df.to_excel("EnrichedAffiliateSheet_Test10.xlsx", index=False)
print("✅ EnrichedAffiliateSheet_Test10.xlsx generated successfully.")
