
import requests
import pandas as pd
from datetime import datetime
import time

# --- Configuration ---
API_KEY = "aja9g1k52mr0co8l8h3csj64io1slb3j7t6se1ic5taar1skedktm3pa95qb5862"
CATEGORY_ID = 14304371  # Beauty
DOMAIN = 5  # Amazon Japan
MAX_ASINS = 5
SAVE_DIR = "C:/Users/Philip Work/Documents/Python"

# --- Helper Functions ---
def fetch_bestsellers(category_id, retries=3):
    url = f"https://api.keepa.com/bestsellers?key={API_KEY}&domain={DOMAIN}&category={category_id}"
    for attempt in range(retries):
        try:
            print(f"Attempt {attempt+1}: Fetching bestsellers...")
            response = requests.get(url, timeout=10)
            response.raise_for_status()
            return response.json().get("asins", [])[:MAX_ASINS]
        except Exception as e:
            print(f"Error fetching bestsellers: {e}")
            time.sleep(2)
    return []

def fetch_product_data(asins, retries=3):
    asin_str = ",".join(asins)
    url = f"https://api.keepa.com/product?key={API_KEY}&domain={DOMAIN}&asin={asin_str}&history=0"
    for attempt in range(retries):
        try:
            print(f"Attempt {attempt+1}: Fetching product data for {asin_str}...")
            response = requests.get(url, timeout=10)
            response.raise_for_status()
            return response.json().get("products", [])
        except Exception as e:
            print(f"Error fetching product data: {e}")
            time.sleep(2)
    return []

# --- Main Execution ---
print("Starting mini live Keepa test...")
try:
    asin_list = fetch_bestsellers(CATEGORY_ID)
    if not asin_list:
        raise ValueError("No ASINs retrieved.")

    products = fetch_product_data(asin_list)
    if not products:
        raise ValueError("No product data retrieved.")

    records = []
    for prod in products:
        records.append({
            "ASIN": prod.get("asin"),
            "Title": prod.get("title", "N/A"),
            "Price (¥)": prod.get("buyBoxPrice", 0) / 100 if prod.get("buyBoxPrice") else "N/A",
            "Link": f"https://www.amazon.co.jp/dp/{prod.get('asin')}"
        })

    df = pd.DataFrame(records)
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
    output_path = f"{SAVE_DIR}/keepa_mini_live_output_{timestamp}.xlsx"
    df.to_excel(output_path, index=False)
    print(f"✅ File saved to: {output_path}")

except Exception as e:
    print(f"❌ Final error: {e}")
