import requests
import time
import pandas as pd

API_KEY = "aja9g1k52mr0co8l8h3csj64io1slb3j7t6se1ic5taar1skedktm3pa95qb5862"
ASINS = [
    "B09Y8H21D2", "B001VZAE9U", "B0DM4KJQQH", "B0D25LBQSV", "B0B973JHTF",
    "B003R6QJHW", "B0D5HGK14R", "B000SJZE4A", "B0D6X7QMVZ", "B001MTCA6K"
]

# Helper to safely extract latest value or return None
def get_last(value):
    if isinstance(value, list):
        return value[-1]
    return value

# Extract product info
def extract_product_info(product):
    stats = product.get("stats", {})
    buy_box_price = get_last(stats.get("buyBoxPrice"))
    referral_fee_pct = product.get("referralFee")

    return {
        "ASIN": product.get("asin"),
        "Title": product.get("title"),
        "Buy Box Price (¥)": buy_box_price / 100 if buy_box_price else None,
        "Referral Fee %": referral_fee_pct,
    }

# Query Keepa
headers = {"Accept-Encoding": "gzip"}
all_results = []
chunk_size = 10  # Only 10 in this test batch

tokens_per_call = 1  # 1 token per ASIN per call for PRODUCT domain
sleep_between_calls = 3  # 3 seconds per 10 tokens to be safe

for i in range(0, len(ASINS), chunk_size):
    chunk = ASINS[i:i + chunk_size]
    print(f"Processing ASINs: {chunk}")

    url = f"https://api.keepa.com/product?key={API_KEY}&domain=5&buybox=1&stats=180&history=0&asin={','.join(chunk)}"
    response = requests.get(url, headers=headers)
    data = response.json()

    if "products" in data:
        for product in data["products"]:
            all_results.append(extract_product_info(product))
    else:
        print(f"Warning: No products returned for chunk {chunk}")

    time.sleep(sleep_between_calls)

# Export to Excel
output_df = pd.DataFrame(all_results)
output_path = "keepa_test_output.xlsx"
output_df.to_excel(output_path, index=False)
print(f"Saved: {output_path}")
