
import pandas as pd
from keepa import Keepa
from tqdm import tqdm
from datetime import datetime
import os

API_KEY = "aja9g1k52mr0co8l8h3csj64io1slb3j7t6se1ic5taar1skedktm3pa95qb5862"
INPUT_FILE = sorted([f for f in os.listdir() if f.startswith("Final_Enriched_Affiliate_Sheet_") and f.endswith(".xlsx")])[-1]
OUTPUT_FILE = f"Final_Enriched_Affiliate_Sheet_{datetime.now().strftime('%Y%m%d_%H%M')}_API.xlsx"

keepa = Keepa(API_KEY)
df = pd.read_excel(INPUT_FILE)
asins = df["ASIN"].dropna().astype(str).tolist()

print("Querying Keepa API...")
products = keepa.query(asins, domain='JP', stats=True)

# Create mapping from ASIN to product data
product_map = {p['asin']: p for p in products if isinstance(p, dict)}

print("Enriching via Keepa API:")
for idx, row in tqdm(df.iterrows(), total=len(df)):
    asin = str(row["ASIN"]).strip()
    product = product_map.get(asin)

    if not product:
        continue

    stats = product.get("stats", {})
    df.at[idx, "isLowest"] = stats.get("isLowest", "")
    df.at[idx, "isLowest90"] = stats.get("isLowest90", "")
    df.at[idx, "buyBox30dDropPercent"] = stats.get("buyBox30dDropPercentage", "")

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

    vc = product.get("variationCount")
    if vc is not None:
        if vc <= 3:
            df.at[idx, "Variation Simplicity"] = "✅ Simple"
        elif vc <= 10:
            df.at[idx, "Variation Simplicity"] = "⚠️ Moderate"
        else:
            df.at[idx, "Variation Simplicity"] = "❌ Complex"
    else:
        df.at[idx, "Variation Simplicity"] = "N/A"

    # Fake Drop Flag
    drop_pct = stats.get("buyBox30dDropPercentage", 0)
    if not stats.get("isLowest90", True) and drop_pct >= 30:
        df.at[idx, "Fake Drop Flag"] = True
    else:
        df.at[idx, "Fake Drop Flag"] = False

    # Grade
    try:
        discount = float(str(row.get("Discount (%)", "")).replace("%", "").strip() or 0)
        sales_rank = int(row.get("Sales Rank", 9999999))
    except:
        discount = 0
        sales_rank = 9999999

    grade = "F"
    if discount >= 50 and sales_rank <= 5000:
        grade = "A"
    elif discount >= 40 and sales_rank <= 10000:
        grade = "B"
    elif discount >= 30 and sales_rank <= 20000:
        grade = "C"
    elif discount >= 20 or sales_rank <= 50000:
        grade = "D"

    if df.at[idx, "Fake Drop Flag"] == True and grade in ["A", "B", "C", "D"]:
        downgrade = {"A": "B", "B": "C", "C": "D", "D": "F"}
        grade = downgrade[grade]

    df.at[idx, "Grade"] = grade

df.to_excel(OUTPUT_FILE, index=False)
print(f"✅ Done. Saved enriched file to {OUTPUT_FILE}")
