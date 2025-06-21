import os
import json
from PIL import Image, ImageDraw, ImageFont

# === CONFIG ===
FOLDER = "C:/Users/Philip Work/Documents/Python"
IMAGE_FOLDER = os.path.join(FOLDER, "affiliate_images")
OUTPUT_FOLDER = os.path.join(FOLDER, "instagram_output")
JSON_PATH = os.path.join(FOLDER, "todays_deals.json")
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# === Load product data ===
with open(JSON_PATH, "r", encoding="utf-8") as f:
    products = json.load(f)

# === Load fonts (ensure these are available or update path) ===
font_title = ImageFont.truetype("arialbd.ttf", 36)
font_text = ImageFont.truetype("arial.ttf", 28)
font_footer = ImageFont.truetype("arialbd.ttf", 30)

# === Loop through each product ===
for p in products:
    asin = p.get("asin")
    if not asin:
        continue

    image_path = os.path.join(IMAGE_FOLDER, f"{asin}.jpg")
    if not os.path.exists(image_path):
        print(f"❌ Image not found for ASIN: {asin}")
        continue

    try:
        # Load and resize product image
        product_img = Image.open(image_path).convert("RGBA").resize((360, 360))

        # Create beige background
        bg_width, bg_height = 768, 1024
        background = Image.new("RGBA", (bg_width, bg_height), (245, 235, 220, 255))
        background.paste(product_img, (60, 80))

        draw = ImageDraw.Draw(background)
        text_x = 450

        # Draw content
        title_lines = p.get("title", "")[:90].replace("\n", " ").split(" ")
        draw.text((text_x, 80), " ".join(title_lines[:5]), font=font_title, fill=(0, 0, 0))

        draw.text((text_x, 240), p.get("stars", ""), font=font_text, fill=(255, 153, 0))
        draw.text((text_x, 275), p.get("discount", ""), font=font_text, fill=(192, 0, 0))
        draw.text((text_x, 320), p.get("price", ""), font=font_text, fill=(0, 0, 0))
        if p.get("coupon"):
            draw.text((text_x, 365), p["coupon"], font=font_text, fill=(255, 204, 0))
        draw.text((text_x, 420), p.get("verified", ""), font=font_text, fill=(100, 100, 100))

        # Footer banner
        footer_text = "ショップのリンクはプロフィールからご覧ください"
        footer_height = 80
        footer_banner = Image.new("RGBA", (bg_width, footer_height), (180, 0, 0, 255))
        draw_footer = ImageDraw.Draw(footer_banner)
        draw_footer.text((100, 20), footer_text, font=font_footer, fill=(255, 255, 255))
        background.paste(footer_banner, (0, bg_height - footer_height))

        # Save output
        output_path = os.path.join(OUTPUT_FOLDER, f"{asin}_insta.png")
        background.save(output_path)
        print(f"✅ Saved: {output_path}")

    except Exception as e:
        print(f"❌ Error creating image for ASIN {asin}: {e}")

print("✅ All previews generated.")
