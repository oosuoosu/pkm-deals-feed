
import requests
import pandas as pd
from datetime import datetime
import time

# --- Configuration ---
API_KEY = "aja9g1k52mr0co8l8h3csj64io1slb3j7t6se1ic5taar1skedktm3pa95qb5862"
DOMAIN = 5  # Amazon Japan
TEST_ASINS = ["B091PLS6MG", "B0DT48NCKC", "B0F5H9F4XL"]
OUTPUT_FOLDER = "C:/Users/Philip Work/Documents/Python"
timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")

print("Starting Keepa ASIN debug test...")

# --- Fetch Product Data ---
def fetch_product_data(asins):
    asin_str = ",".join(asins)
    url = f"https://api.keepa.com/product?key={API_KEY}&domain={DOMAIN}&asin={asin_str}&history=0"
    print(f"Fetching product data from: {url}")
    try:
        response = requests.get(url)
        print(f"HTTP Status Code: {response.status_code}")
        response.raise_for_status()
        print("✅ Product data retrieved successfully.")
        return response.json().get("products", [])
    except Exception as e:
        print(f"❌ Error during product fetch: {e}")
        return []

try:
    products = fetch_product_data(TEST_ASINS)
    if not products:
        print("⚠️ No products returned. Exiting.")
    else:
        records = []
        for prod in products:
            records.append({
                "ASIN": prod.get("asin"),
                "Title": prod.get("title", "N/A"),
                "Price (¥)": prod.get("buyBoxPrice", 0) / 100 if prod.get("buyBoxPrice") else "N/A",
                "Link": f"https://www.amazon.co.jp/dp/{prod.get('asin')}"
            })

        df = pd.DataFrame(records)
        output_path = f"{OUTPUT_FOLDER}/keepa_asin_debug_output_{timestamp}.xlsx"
        df.to_excel(output_path, index=False)
        print(f"✅ Excel saved to: {output_path}")
except Exception as e:
    print(f"❌ Final error: {e}")
