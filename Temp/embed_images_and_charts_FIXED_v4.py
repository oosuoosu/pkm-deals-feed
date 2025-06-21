from datetime import datetime
from io import BytesIO
import pandas as pd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
import requests
from tqdm import tqdm

# === CONFIG ===
INPUT_FILE = "Final_Enriched_Affiliate_Sheet_Styled_20250611_1127.xlsx"
OUTPUT_FILE = f"Final_Enriched_Affiliate_Sheet_ImagesOnly_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

# === LOAD ===
df = pd.read_excel(INPUT_FILE)
wb = load_workbook(INPUT_FILE)
ws = wb.active

# === EMBED PRODUCT IMAGES ONLY ===
print("🖼️ Embedding product images (no charts)...")
for idx, row in tqdm(df.iterrows(), total=len(df), desc="Embedding"):
    asin = row.get("ASIN", "")
    image_url = row.get("Image", "")
    if not asin or pd.isna(asin):
        continue

    if image_url and not pd.isna(image_url):
        try:
            img_data = requests.get(image_url, timeout=5).content
            img = XLImage(BytesIO(img_data))
            img.width, img.height = 80, 80
            ws.row_dimensions[idx + 2].height = 100
            ws.add_image(img, f"A{idx + 2}")
        except Exception as e:
            print(f"⚠️ Image error for row {idx}: {e}")

# === SAVE ===
wb.save(OUTPUT_FILE)
print(f"✅ Saved to: {OUTPUT_FILE}")
