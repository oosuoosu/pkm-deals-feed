
import pandas as pd
import requests
import time
from tqdm import tqdm
from datetime import datetime
import pytz

# === CONFIGURATION ===
API_KEY = "aja9g1k52mr0co8l8h3csj64io1slb3j7t6se1ic5taar1skedktm3pa95qb5862"
INPUT_FILE = "KeepaExport-2025-ProductFinder.xlsx"
TEMPLATE_FILE = "Final_Enriched_Affiliate_Template.xlsx"
STARTING_TOKENS = 1260
TOKENS_PER_REQUEST = 10
RECOVERY_INTERVAL_SEC = 3
TIMEZONE = "America/Vancouver"

def fetch_keepa_data(asin):
    url = (
        f"https://api.keepa.com/product?key={API_KEY}&domain=5"
        f"&buybox=1&stats=180&history=0&asin={asin}"
    )
    response = requests.get(url)
    if response.status_code == 200:
        data = response.json().get("products", [])
        return data[0] if data else None
    else:
        print(f"API Error for {asin}: {response.text}")
        return None

def get_discount_percent(product):
    try:
        current = product["stats"].get("buyBoxPrice", [None])[-1]
        avg_90 = product["stats"].get("avg90", [None])[-1]
        if current and avg_90 and avg_90 != 0:
            return round((avg_90 - current) / avg_90 * 100, 2)
    except:
        pass
    return None

def calculate_grade(discount, rank):
    if discount is None or rank is None:
        return "F"
    if discount >= 50 and rank <= 5000:
        return "A"
    elif discount >= 40 and rank <= 10000:
        return "B"
    elif discount >= 30 and rank <= 20000:
        return "C"
    elif discount >= 20 or rank <= 50000:
        return "D"
    return "F"

def calculate_fake_drop(product):
    try:
        is_lowest = product.get("isLowest", False)
        is_lowest90 = product.get("isLowest90", False)
        drop_percent = product.get("buyBox30dDropPercent", 0)
        return not is_lowest90 and drop_percent >= 30
    except:
        return False

def enrich_row(row, product):
    stats = product.get("stats", {})
    row["Discount (%)"] = get_discount_percent(product)
    row["isLowest"] = product.get("isLowest", False)
    row["isLowest90"] = product.get("isLowest90", False)
    row["buyBox30dDropPercent"] = product.get("buyBox30dDropPercent", None)
    row["Grade"] = calculate_grade(row["Discount (%)"], row["Sales Rank"])
    row["Fake Drop Flag"] = calculate_fake_drop(product)
    row["Affiliate Link"] = f"https://www.amazon.co.jp/dp/{row['ASIN']}?tag=pkmsalechanne-22"
    return row

# === MAIN PROCESS ===
template_df = pd.read_excel(TEMPLATE_FILE)
input_df = pd.read_excel(INPUT_FILE)
asin_list = input_df["ASIN"].dropna().astype(str).tolist()

merged_df = template_df.copy()
merged_df["ASIN"] = asin_list  # Inject ASINs from input

remaining_tokens = STARTING_TOKENS

for i, asin in enumerate(tqdm(asin_list, desc="Processing ASINs")):
    if remaining_tokens < TOKENS_PER_REQUEST:
        while remaining_tokens < TOKENS_PER_REQUEST:
            time.sleep(RECOVERY_INTERVAL_SEC)
            remaining_tokens += 1
    product = fetch_keepa_data(asin)
    remaining_tokens -= TOKENS_PER_REQUEST
    if product:
        row = merged_df.iloc[i]
        merged_df.iloc[i] = enrich_row(row, product)

# === SAVE OUTPUT ===
timestamp = datetime.now(pytz.timezone(TIMEZONE)).strftime("%Y%m%d_%H%M")
output_path = f"Final_Enriched_Affiliate_Sheet_{timestamp}.xlsx"
merged_df.to_excel(output_path, index=False)
print(f"✅ Done. File saved as {output_path}")
