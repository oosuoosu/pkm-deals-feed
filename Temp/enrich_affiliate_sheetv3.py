# Fix script generation by using double braces for escaped curly braces in f-strings
from datetime import datetime

script_path = "/mnt/data/enrich_affiliate_sheet_corrected.py"

template_columns_code = repr(template_columns)

script_code = f"""
import pandas as pd
import requests
import time
from tqdm import tqdm

API_KEY = "aja9g1k52mr0co8l8h3csj64io1slb3j7t6se1ic5taar1skedktm3pa95qb5862"
INPUT_FILE = "KeepaExport-2025-ProductFinder.xlsx"
TEMPLATE_FILE = "Final_Enriched_Affiliate_Template.xlsx"
OUTPUT_FILE = "Final_Enriched_Affiliate_Sheet_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

STARTING_TOKENS = 1260
TOKENS_PER_REQUEST = 10
RECOVERY_INTERVAL_SEC = 3

# Required columns in the final output
FINAL_COLUMNS = {template_columns_code}

def generate_affiliate_link(asin):
    return f"https://www.amazon.co.jp/dp/{{asin}}?tag=pkmsalechanne-22"

def load_product_finder(path):
    df = pd.read_excel(path)
    df['ASIN'] = df['ASIN'].astype(str)
    return df

def fetch_keepa_data(asin):
    url = f"https://api.keepa.com/product?key={{API_KEY}}&domain=5&buybox=1&stats=180&history=0&asin={{asin}}"
    response = requests.get(url)
    if response.status_code == 200:
        return response.json().get("products", [])[0]
    return None

def calculate_grade(discount, sales_rank):
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
    else:
        return "F"

def check_fake_drop(product):
    drop = product.get("stats", {{}}).get("buyBox30dDropPercent", 0)
    is_lowest = product.get("stats", {{}}).get("isLowest", False)
    is_lowest90 = product.get("stats", {{}}).get("isLowest90", False)
    return not is_lowest and not is_lowest90 and drop >= 30

def enrich_data(base_df):
    enriched_rows = []
    tokens = STARTING_TOKENS

    for _, row in tqdm(base_df.iterrows(), total=len(base_df), desc="Processing"):
        asin = row.get("ASIN")
        result = {{col: None for col in FINAL_COLUMNS}}
        result["ASIN"] = asin
        result["Title"] = row.get("Title")
        result["URL: Amazon"] = row.get("URL: Amazon")
        result["Sales Rank"] = row.get("Sales Rank: Current")
        result["Rating"] = row.get("Reviews: Rating")
        result["Review Count"] = row.get("Reviews: Review Count")
        result["Price (¥)"] = row.get("Buy Box: Current")
        result["Referral Fee %"] = row.get("Referral Fee %")
        result["Referral Fee based on current Buy Box price"] = row.get("Referral Fee based on current Buy Box price")
        result["Buy Box: Stock"] = row.get("Buy Box: Stock")
        result["Affiliate Link"] = generate_affiliate_link(asin)

        if tokens < TOKENS_PER_REQUEST:
            while tokens < TOKENS_PER_REQUEST:
                time.sleep(RECOVERY_INTERVAL_SEC)
                tokens += 1

        product = fetch_keepa_data(asin)
        tokens -= TOKENS_PER_REQUEST

        if product:
            stats = product.get("stats", {{}})
            result["Discount (%)"] = stats.get("buyBox30dDropPercent")
            result["isLowest"] = stats.get("isLowest")
            result["isLowest90"] = stats.get("isLowest90")
            result["buyBox30dDropPercent"] = stats.get("buyBox30dDropPercent")
            result["Fake Drop Flag"] = check_fake_drop(product)
            result["Grade"] = calculate_grade(result["Discount (%)"], result["Sales Rank"])
            result["Commission Tier"] = (
                "Tier 1" if row.get("Referral Fee %", 0) >= 12 else
                "Tier 2" if row.get("Referral Fee %", 0) >= 8 else
                "Tier 3" if row.get("Referral Fee %", 0) >= 5 else "Tier 4"
            )

        enriched_rows.append(result)

    return pd.DataFrame(enriched_rows, columns=FINAL_COLUMNS)

def main():
    base_df = load_product_finder(INPUT_FILE)
    result_df = enrich_data(base_df)
    result_df.to_excel(OUTPUT_FILE, index=False)
    print(f"✅ Done. File saved to {{OUTPUT_FILE}}")

if __name__ == "__main__":
    main()
"""

with open(script_path, "w", encoding="utf-8") as f:
    f.write(script_code)

script_path
