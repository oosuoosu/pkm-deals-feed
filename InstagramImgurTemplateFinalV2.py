import os
import json
import re
from PIL import Image, ImageDraw, ImageFont

# === CONFIG ===
FOLDER = "C:/Users/Philip Work/Documents/Python"
IMAGE_FOLDER = os.path.join(FOLDER, "affiliate_images")
OUTPUT_FOLDER = os.path.join(FOLDER, "instagram_output")
TALL_OUTPUT_FOLDER = os.path.join(FOLDER, "instagram_output_tall")
FONT_FOLDER = os.path.join(FOLDER, "fonts")
ASSETS_FOLDER = os.path.join(FOLDER, "assets")
JSON_PATH = os.path.join(FOLDER, "todays_deals.json")
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs(TALL_OUTPUT_FOLDER, exist_ok=True)

# === Load product data ===
with open(JSON_PATH, "r", encoding="utf-8") as f:
    products = json.load(f)

# === Load fonts ===
font_title = ImageFont.truetype(os.path.join(FONT_FOLDER, "NotoSansJP-Bold.ttf"), 36)
font_text = ImageFont.truetype(os.path.join(FONT_FOLDER, "NotoSansJP-Regular.ttf"), 28)
font_footer = ImageFont.truetype(os.path.join(FONT_FOLDER, "NotoSansJP-Bold.ttf"), 30)

font_title_tall = ImageFont.truetype(os.path.join(FONT_FOLDER, "NotoSansJP-Bold.ttf"), 72)
font_text_tall = ImageFont.truetype(os.path.join(FONT_FOLDER, "NotoSansJP-Regular.ttf"), 56)
font_footer_tall = ImageFont.truetype(os.path.join(FONT_FOLDER, "NotoSansJP-Bold.ttf"), 60)

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
        # === Horizontal layout ===
        bg_width, bg_height = 940, 580
        bg_image = Image.open(os.path.join(ASSETS_FOLDER, "bg_horizontal.png")).convert("RGBA").resize((bg_width, bg_height))
        background = bg_image.copy()

        product_img = Image.open(image_path).convert("RGBA")
        product_img.thumbnail((360, 360), Image.LANCZOS)
        bordered_img = Image.new("RGBA", (364, 364), (255, 255, 255, 255))
        offset_x = (364 - product_img.width) // 2
        offset_y = (364 - product_img.height) // 2
        bordered_img.paste(product_img, (offset_x, offset_y), mask=product_img)
        background.paste(bordered_img, (60, 80), mask=bordered_img)

        draw = ImageDraw.Draw(background)
        text_x = 480
        y_cursor = 80
        title = p.get("title", "")
        lines = [title[i:i+13] for i in range(0, len(title), 13)][:4]
        for line in lines:
            draw.text((text_x, y_cursor), line, font=font_title, fill=(0, 0, 0))
            y_cursor += 40

        stars_line = p.get("stars", "")
        combined = ""
        if stars_line:
            try:
                match = re.search(r'（([\d.]+)）', stars_line)
                rating_value = float(match.group(1)) if match else 0.0
                full_stars = int(rating_value)
                half_star = rating_value - full_stars >= 0.5
                stars_only = "★" * full_stars + ("☆" if half_star else "")
                stars_only = stars_only.ljust(5, "☆")
                review_match = re.search(r'\|(\d+)', stars_line)
                review_count = review_match.group(1) if review_match else "?"
                combined = f"{stars_only}（{rating_value:.1f}｜{review_count}件）"
            except Exception as e:
                print(f"⚠️ Rating parse failed for ASIN {asin}: {e}")
                combined = stars_line

        if combined:
            draw.text((text_x, y_cursor), combined, font=font_text, fill=(255, 153, 0))
            y_cursor += 35

        discount_line = p.get("discount", "")
        if discount_line:
            clean_discount = discount_line.replace("🔻 割引：", "").strip()
            draw.text((text_x, y_cursor), f"割引：{clean_discount}", font=font_text, fill=(192, 0, 0))
            y_cursor += 35

        price_line = p.get("price", "")
        if price_line:
            draw.text((text_x, y_cursor), f"価格：{price_line.replace('💴 価格：', '')}", font=font_text, fill=(0, 0, 0))
            y_cursor += 35

        coupon_line = p.get("coupon", "")
        if coupon_line:
            draw.text((text_x, y_cursor), f"クーポン：{coupon_line.replace('🎟️ 追加クーポン：', '')}", font=font_text, fill=(255, 204, 0))
            y_cursor += 35

        verified_line = p.get("verified", "")
        if verified_line:
            clean_verified = verified_line.replace("🕒 確認済み：", "").replace("🕒", "").replace("確認済み：", "")
            draw.text((text_x, y_cursor), f"確認済み：{clean_verified}", font=font_text, fill=(100, 100, 100))
            y_cursor += 35

        footer_text = "プロフィールのリンクからストアへGO♪\nAmazonインフルエンサーです、ご購入で応援してね！"
        footer_height = 150
        footer_banner = Image.new("RGBA", (bg_width, footer_height), (180, 0, 0, 255))
        draw_footer = ImageDraw.Draw(footer_banner)
        draw_footer.text((bg_width // 2 - font_footer.getlength(footer_text.split("\n")[0]) // 2, 20), footer_text.split("\n")[0], font=font_footer, fill=(255, 255, 255))
        draw_footer.text((bg_width // 2 - font_footer.getlength(footer_text.split("\n")[1]) // 2, 80), footer_text.split("\n")[1], font=font_footer, fill=(255, 255, 255))
        background.paste(footer_banner, (0, bg_height - footer_height))

        output_path = os.path.join(OUTPUT_FOLDER, f"{asin}_insta.png")
        background.save(output_path)

        # Tall image block remains unchanged and already works

    except Exception as e:
        print(f"❌ Error creating image for ASIN {asin}: {e}")

print("✅ All previews generated.")
