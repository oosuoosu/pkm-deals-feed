import requests
import pandas as pd
from datetime import datetime
import os

# --- Configuration ---
API_KEY = "aja9g1k52mr0co8l8h3csj64io1slb3j7t6se1ic5taar1skedktm3pa95qb5862"
CATEGORY_ID = 14304371  # Beauty
DOMAIN = 5  # Amazon Japan
MAX_ASINS = 5  # Limit for mini version
SAVE_DIR = "C:/Users/Philip Work/Documents/Python"

# --- Helper Functions ---
def fetch_bestsellers(category_id):
    url = f"https://api.keepa.com/bestsellers?key={API_KEY}&domain={DOMAIN}&category={category_id}"
    response = requests.get(url)
    response.raise_for_status()
    return response.json().get("asins", [])[:MAX_ASINS]

def fetch_product_data(asins):
    asin_str = ",".join(asins)
    url = f"https://api.keepa.com/product?key={API_KEY}&domain={DOMAIN}&asin={asin_str}&history=0"
    response = requests.get(url)
    response.raise_for_status()
    return response.json().get("products", [])

# --- Main Execution ---
print("Fetching bestsellers for Beauty...")
try:
    asin_list = fetch_bestsellers(CATEGORY_ID)
    products = fetch_product_data(asin_list)

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
    output_path = os.path.join(SAVE_DIR, f"keepa_mini_live_output_{timestamp}.xlsx")
    df.to_excel(output_path, index=False)
    print(f"Saved to {output_path}")

except Exception as e:
    print(f"Error: {e}")
