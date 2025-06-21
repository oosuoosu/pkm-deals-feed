import pandas as pd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
import requests
from io import BytesIO
import matplotlib.pyplot as plt
import os
from tqdm import tqdm
from keepa import Keepa
from datetime import datetime

# === CONFIG ===
INPUT_FILE = "Final_Enriched_Affiliate_Sheet_Styled_20250611_1127.xlsx"
OUTPUT_FILE = f"Final_Enriched_Affiliate_Sheet_ImagesCharts_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
API_KEY = "aja9g1k52mr0co8l8h3csj64io1slb3j7t6se1ic5taar1skedktm3pa95qb5862"
DOMAIN = "JP"

# === INITIALIZE ===
keepa = Keepa(API_KEY)
df = pd.read_excel(INPUT_FILE)
wb = load_workbook(INPUT_FILE)
ws = wb.active

# === HELPER: Generate and return chart image from Keepa CSV ===
def generate_chart(asin):
    try:
        product = keepa.query([asin], domain=DOMAIN, history=True)[0]
        if "csv" not in product or len(product["csv"][0]) == 0:
            return None

        price_data = product["csv"][0]  # Amazon price
        if not price_data:
            return None

        timestamps = [point[0] for point in price_data]
        prices = [point[1] / 100 for point in price_data]

        if not prices:
            return None

        plt.figure(figsize=(4, 2))
        plt.plot(prices, linewidth=1.5)
        plt.title("Price Trend (2mo)")
        plt.xlabel("Time")
        plt.ylabel("¥ Price")
        plt.grid(True)
        plt.tight_layout()

        buf = BytesIO()
        plt.savefig(buf, format="png", dpi=100)
        plt.close()
        buf.seek(0)
        return buf

    except Exception as e:
        print(f"⚠️ Chart error for ASIN {asin}: {e}")
        return None

# === PROCESS EACH ROW ===
print("🖼️ Embedding Images and Charts...")
for idx, row in tqdm(df.iterrows(), total=len(df), desc="Embedding"):
    asin = row.get("ASIN", "")
    if not asin or pd.isna(asin):
        continue

    # Embed Image
    image_url = row.get("Image", "")
    if image_url and not pd.isna(image_url):
        try:
            img_data = requests.get(image_url, timeout=5).content
            img = XLImage(BytesIO(img_data))
            img.width, img.height = 80, 80
            ws.row_dimensions[idx + 2].height = 80  # Row height
            ws.add_image(img, f"A{idx + 2}")
        except Exception as e:
            print(f"⚠️ Image error for row {idx}: {e}")

    # Embed Chart
    chart_img = generate_chart(asin)
    if chart_img:
        try:
            chart = XLImage(chart_img)
            chart.width, chart.height = 120, 80
            ws.add_image(chart, f"B{idx + 2}")
        except Exception as e:
            print(f"⚠️ Chart insert error for row {idx}: {e}")

# === SAVE OUTPUT ===
wb.save(OUTPUT_FILE)
print(f"✅ Done! Saved to: {OUTPUT_FILE}")
