from datetime import datetime
import pandas as pd
from keepa import Keepa
from tqdm import tqdm

# --- Configuration ---
API_KEY = "aja9g1k52mr0co8l8h3csj64io1slb3j7t6se1ic5taar1skedktm3pa95qb5862"
INPUT_FILE = "Final_Enriched_Affiliate_Sheet_20250610_2242.xlsx"  # <-- replace as needed
OUTPUT_FILE = f"Final_Enriched_Affiliate_Sheet_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
TRACKING_ID = "pkmsalechanne-22"

# --- Connect to Keepa ---
keepa = Keepa(API_KEY)

# --- Load Data ---
df = pd.read_excel(INPUT_FILE)
asins = df["ASIN"].dropna().astype(str).tolist()

# --- Query Keepa ---
print("Querying Keepa API...")
products = keepa.query(asins, domain='JP', stats=True)

# --- Enrich Rows ---
print("Enriching via Keepa API...")
for idx, asin in enumerate(tqdm(asins, desc="Enriching ASINs")):
    product = next((p for p in products if p.get("asin") == asin), {})
    stats = product.get("stats", {})

    def safe_bool(val):
        if isinstance(val, list):
            return bool(val[-1])
        return bool(val)

    def safe_float(val):
        if isinstance(val, list):
            return float(val[-1])
        return float(val) if val is not None else ""

    # Assign API-derived fields
    df.at[idx, "isLowest"] = safe_bool(stats.get("isLowest", False))
    df.at[idx, "isLowest90"] = safe_bool(stats.get("isLowest90", False))
    df.at[idx, "buyBox30dDropPercent"] = safe_float(stats.get("buyBox30dDropPercentage"))

    # Trending BSR
    sr30 = stats.get("salesRank30")
    sr90 = stats.get("salesRank90")
    if sr30 and sr90:
        if sr30 < sr90:
            df.at[idx, "Trending BSR"] = "↑ Improving"
        elif sr30 > sr90:
            df.at[idx, "Trending BSR"] = "↓ Worsening"
        else:
            df.at[idx, "Trending BSR"] = "→ Stable"
    else:
        df.at[idx, "Trending BSR"] = "N/A"

    # Variation Simplicity
    var_count = product.get("variationCount")
    if var_count is not None:
        if var_count <= 3:
            df.at[idx, "Variation Simplicity"] = "✅ Simple"
        elif var_count <= 10:
            df.at[idx, "Variation Simplicity"] = "⚠️ Moderate"
        else:
            df.at[idx, "Variation Simplicity"] = "❌ Complex"
    else:
        df.at[idx, "Variation Simplicity"] = "N/A"

    # Fake Drop Flag
    is_lowest90 = safe_bool(stats.get("isLowest90", False))
    drop_pct = safe_float(stats.get("buyBox30dDropPercentage"))
    fake_drop = not is_lowest90 and drop_pct >= 30 if drop_pct != "" else False
    df.at[idx, "Fake Drop Flag"] = fake_drop

    # Grade (preserve Discount column)
    discount = df.at[idx, "Discount (%)"]
    rank = df.at[idx, "Sales Rank"]
    grade = "F"
    if pd.notna(discount) and pd.notna(rank):
        try:
            discount = float(discount)
            rank = int(rank)
            if discount >= 50 and rank <= 5000:
                grade = "A"
            elif discount >= 40 and rank <= 10000:
                grade = "B"
            elif discount >= 30 and rank <= 20000:
                grade = "C"
            elif discount >= 20 or rank <= 50000:
                grade = "D"
        except:
            grade = "F"
    if fake_drop and grade in "ABCD":
        downgrade = {"A": "B", "B": "C", "C": "D", "D": "F"}
        grade = downgrade[grade]
    df.at[idx, "Grade"] = grade

    # Affiliate Link
    df.at[idx, "Affiliate Link"] = f"https://www.amazon.co.jp/dp/{asin}/?tag={TRACKING_ID}"

# --- Save Output ---
df.to_excel(OUTPUT_FILE, index=False)
print(f"✅ Saved enriched file with API data to {OUTPUT_FILE}")
