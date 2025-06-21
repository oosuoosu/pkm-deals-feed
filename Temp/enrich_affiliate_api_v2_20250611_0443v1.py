
import pandas as pd
import requests
import time
from tqdm import tqdm

API_KEY = "aja9g1k52mr0co8l8h3csj64io1slb3j7t6se1ic5taar1skedktm3pa95qb5862"
INPUT_FILE = "Final_Enriched_Affiliate_Sheet_20250611_0443.xlsx"
OUTPUT_FILE = "Final_Enriched_Affiliate_Sheet_Keepa_API_20250611_0443.xlsx"

START_TOKENS = 1260
TOKEN_RECOVERY_RATE = 21
TOKENS_PER_ASIN = 10
SLEEP_INTERVAL = 60 / TOKEN_RECOVERY_RATE

def calculate_grade(discount, sales_rank, fake_drop):
    if discount is None or sales_rank is None:
        return "F"
    if discount >= 50 and sales_rank <= 5000:
        grade = "A"
    elif discount >= 40 and sales_rank <= 10000:
        grade = "B"
    elif discount >= 30 and sales_rank <= 20000:
        grade = "C"
    elif discount >= 20 or sales_rank <= 50000:
        grade = "D"
    else:
        grade = "F"
    return chr(ord(grade) + 1) if fake_drop and grade < "F" else grade

def calculate_fake_drop(is_lowest, is_lowest_90, drop_percent):
    return not is_lowest and not is_lowest_90 and drop_percent >= 30

def calculate_trending_bsr(sr30, sr90):
    if sr30 is not None and sr90 is not None:
        if sr30 < sr90:
            return "↑ Improving"
        elif sr30 > sr90:
            return "↓ Worsening"
        return "→ Stable"
    return "N/A"

def calculate_variation_simplicity(vc):
    if vc is None:
        return "N/A"
    if vc <= 3:
        return "✅ Simple"
    elif vc <= 10:
        return "⚠️ Moderate"
    return "❌ Complex"

def fetch_keepa_data(asins):
    url = f"https://api.keepa.com/product?key={API_KEY}&domain=5&stats=180&buybox=1&history=0&asin={','.join(asins)}"
    r = requests.get(url)
    return r.json().get("products", [])

def main():
    df = pd.read_excel(INPUT_FILE)
    tokens = START_TOKENS

    for i in tqdm(range(0, len(df), 10), desc="Enriching via Keepa API"):
        batch = df.iloc[i:i+10]
        asins = batch["ASIN"].dropna().astype(str).tolist()

        if tokens < TOKENS_PER_ASIN * len(asins):
            time.sleep(SLEEP_INTERVAL)
            tokens += TOKEN_RECOVERY_RATE

        tokens -= TOKENS_PER_ASIN * len(asins)
        products = fetch_keepa_data(asins)

        for product in products:
            asin = product.get("asin")
            stats = product.get("stats", {})
            idx_list = df[df["ASIN"] == asin].index.tolist()
            if not idx_list:
                continue
            idx = idx_list[0]

            discount = stats.get("buyBox30dDropPercentage")
            is_lowest = stats.get("isLowest", False)
            is_lowest_90 = stats.get("isLowest90", False)
            sr30 = stats.get("salesRank30")
            sr90 = stats.get("salesRank90")
            vc = product.get("variationCount")

            fake_drop = calculate_fake_drop(is_lowest, is_lowest_90, discount if discount else 0)
            grade = calculate_grade(discount, df.at[idx, "Sales Rank"], fake_drop)
            bsr = calculate_trending_bsr(sr30, sr90)
            variation = calculate_variation_simplicity(vc)

            df.at[idx, "isLowest"] = is_lowest
            df.at[idx, "isLowest90"] = is_lowest_90
            df.at[idx, "buyBox30dDropPercent"] = discount
            df.at[idx, "Fake Drop Flag"] = fake_drop
            df.at[idx, "Grade"] = grade
            df.at[idx, "Trending BSR"] = bsr
            df.at[idx, "Variation Simplicity"] = variation

    df.to_excel(OUTPUT_FILE, index=False)
    print(f"✅ Saved enriched data to {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
