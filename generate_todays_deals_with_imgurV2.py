import os
import pandas as pd
import json
import requests
from datetime import datetime

# === CONFIG ===
FOLDER = "C:/Users/Philip Work/Documents/Python"
IMAGE_FOLDER = os.path.join(FOLDER, "affiliate_images")
PREFIX = "Final_Enriched_Affiliate_ImagesOnly_"
OUTPUT_JSON = os.path.join(FOLDER, "todays_deals.json")
IMGUR_CLIENT_ID = "c296b8f8ec0baa0"
AFFILIATE_TAG = "pkmsalechanne-22"

# === Upload Image to Imgur ===
def upload_to_imgur(image_path):
    try:
        headers = {"Authorization": f"Client-ID {IMGUR_CLIENT_ID}"}
        with open(image_path, "rb") as img_file:
            response = requests.post("https://api.imgur.com/3/image", headers=headers, files={"image": img_file})
        if response.status_code == 200:
            return response.json()["data"]["link"]
        else:
            print(f"❌ Failed to upload {image_path}: {response.text}")
            return None
    except Exception as e:
        print(f"❌ Exception uploading {image_path}: {e}")
        return None

# === Locate Latest Excel ===
files = [f for f in os.listdir(FOLDER) if f.startswith(PREFIX) and f.endswith(".xlsx")]
if not files:
    raise FileNotFoundError(f"No Excel file found in {FOLDER} with prefix '{PREFIX}'")

latest_file = max(files, key=lambda f: os.path.getmtime(os.path.join(FOLDER, f)))
excel_path = os.path.join(FOLDER, latest_file)
print(f"📂 Using input file: {excel_path}")

# === Load Excel and Filter ===
df = pd.read_excel(excel_path)
df = df[df["Marketing Approved"].astype(str).str.lower().str.strip() == "yes"]

# === Format Each Product ===
products = []
for _, row in df.iterrows():
    asin = row["ASIN"]
    title = str(row.get("Title", "")).strip()
    rating = row.get("Rating", "")
    review_count = row.get("Review Count", "")
    discount = row.get("Discount (%)", "")
    price = row.get("New Price", "")
    coupon = row.get("One Time Coupon: Absolute", "")
    verified = row.get("Verified", "")
    amazon_url = f"https://www.amazon.co.jp/dp/{asin}?tag={AFFILIATE_TAG}"

    stars = ""
    try:
        stars = "⭐" * int(float(rating)) + ("✴︎" if float(rating) % 1 >= 0.5 else "")
    except:
        pass

    discount_str = f"{int(round(float(discount) * 100))}% OFF" if pd.notna(discount) else ""
    price_str = f"¥{int(price):,}" if pd.notna(price) else "価格不明"
    coupon_str = f"¥{int(coupon)}" if pd.notna(coupon) and float(coupon) > 0 else None

    image_path = os.path.join(IMAGE_FOLDER, f"{asin}.jpg")
    imgur_link = upload_to_imgur(image_path) if os.path.exists(image_path) else None

    product_block = {
        "asin": asin,
        "title": title,
        "stars": f"{stars}（{rating}）｜{int(review_count)}件" if pd.notna(rating) and pd.notna(review_count) else "",
        "discount": f"🔻 割引：{discount_str}" if discount_str else "",
        "price": f"💴 価格：{price_str}",
        "coupon": f"🎟️ 追加クーポン：{coupon_str}" if coupon_str else "",
        "verified": f"🕒 確認済み：{verified}",
        "link": "👉 Amazonで見る（アフィリエイトリンク）",
        "url": amazon_url,
        "image": imgur_link
    }

    products.append(product_block)

# === Write JSON Output ===
with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
    json.dump(products, f, ensure_ascii=False, indent=2)

print(f"✅ JSON with affiliate-tagged Imgur links saved to: {OUTPUT_JSON}")
