from datetime import datetime
from io import BytesIO
import matplotlib.pyplot as plt
import pandas as pd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
import requests
from tqdm import tqdm
from keepa import Keepa
import os

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

# === CHART GENERATOR ===
def generate_chart(asin):
    try:
        product = keepa.query([asin], domain=DOMAIN, history=True)[0]
        buy_box_data = product.get("csv", {}).get(10)

        if not buy_box_data or not isinstance(buy_box_data, list):
            raise ValueError("Invalid or missing Buy Box data")

        # Keep only last 864 data points (~3 months at 5-min intervals)
        recent = buy_box_data[-864:]
        prices = [p[1] / 100 for p in recent if isinstance(p, list) and p[1] > 0]

        if not prices:
            raise ValueError("No valid Buy Box prices")

        times = list(range(len(prices)))

        plt.figure(figsize=(4, 2))
        plt.plot(times, prices, linewidth=1.5, color='green')
        plt.title("Buy Box Price Trend (3 mo)")
        plt.ylabel("¥")
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

# === EMBED IMAGES + CHARTS ===
print("🖼️ Embedding Images and Charts...")
for idx, row in tqdm(df.iterrows(), total=len(df), desc="Embedding"):
    asin = row.get("ASIN", "")
    if not asin or pd.isna(asin):
        continue

    # Embed Product Image
    image_url = row.get("Image", "")
    if image_url and not pd.isna(image_url):
        try:
            img_data = requests.get(image_url, timeout=5).content
            img = XLImage(BytesIO(img_data))
            img.width, img.height = 80, 80
            ws.row_dimensions[idx + 2].height = 100
            ws.add_image(img, f"A{idx + 2}")
        except Exception as e:
            print(f"⚠️ Image error for row {idx}: {e}")

    # Embed Keepa Chart
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
